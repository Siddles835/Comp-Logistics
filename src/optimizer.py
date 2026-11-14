from pathlib import Path
import pandas as pd
import pulp
from openpyxl import load_workbook

here = Path(__file__).resolve().parent
data = here / "data"
if not data.exists():
    data = here.parent / "data"
    
output = data.parent / "outputs"
output.mkdir(exist_ok=True)

tenth = 10

def find_file(name):
    for file in data.glob("*.xlsx"):
        if name.lower() in file.name.lower():
            if not file.name.startswith("~$"):
                return file
    raise FileNotFoundError("Couldn't find:", name)

def load_file(name):
    file = find_file(name)
    table = pd.read_excel(file)
    table.columns = [column.strip() if isinstance(column, str) else column for column in table.columns]
    
    for column in table.columns:
        if table[column].dtype == object:
            table[column] = table[column].map(lambda value: value.strip() if isinstance(value, str) else value)
    return table

suppliers_data = load_file("suppliers")
plants_data = load_file("processors")
distributors_data = load_file("distributors")
requirements_data = load_file("requirements")
shipping_data = load_file("shipping")
distances_data = load_file("distance")

suppliers = ["Pacific Alloys", "Midwest Makers", "Eastern Metal"]
plants = plants_data["Manufacturing Plant"].tolist()
distributors = distributors_data["Distributors"].tolist()
materials = suppliers_data["Material"].tolist()
products = ["Whos-It (WHO)", "What-Nots (WHAT)", "Thing-a-Ma-Bob (THING)"]
product_codes = ["WHO", "WHAT", "THING"]

material_cost = {}
material_time = {}

for index, row in suppliers_data.iterrows():
    material = row["Material"]
    for supplier in suppliers:
        material_cost[supplier, material] = float(row["Cost (" + supplier + ")"])
        material_time[supplier, material] = float(row["Max Monthly (" + supplier + ")"])

processing_time = {}
hourly_wage = {}
plant_capacity = {}

for index, row in plants_data.iterrows():
    plant = row["Manufacturing Plant"]
    hourly_wage[plant] = float(row["Cost per Hour"])
    plant_capacity[plant] = float(row["Max Capacity (hours)"])
    for product, code in zip(products, product_codes):
        processing_time[plant, product] = float(row["Processing Time, Hours (" + code + ")"])

material_needed = {}
for index, row in requirements_data.iterrows():
    product = row["Product"]
    for material in materials:
        material_needed[product, material] = float(row[material])

market_demand = {}
selling_price = {}

for index, row in distributors_data.iterrows():
    distributor = row["Distributors"]
    for product, code in zip(products, product_codes):
        market_demand[distributor, product] = float(row["Max Market Demand (" + code + ")"])
        selling_price[distributor, product] = float(row["Selling Price (" + code + ")"])

print("SHIPPING COLUMNS:", shipping_data.columns.tolist())
print(shipping_data.head())
exit()

distance = {}
for index, row in distances_data.iterrows():
    starting_location = row["Location"]
    for destination in distances_data.columns[1:]:
        if pd.notna(row[destination]):
            distance[starting_location, destination] = float(row[destination])

def get_distance(start, end):
    if (start, end) not in distance:
        raise ValueError("Missing distance from: {} to {}".format(start, end))
    return distance[start, end]

def to_tenths(quantity):
    return int(round(float(quantity) * tenth))

model = pulp.LpProblem("Supply_Chain_Optimization", pulp.LpMaximize)
make = pulp.LpVariable.dicts("Make", (plants, products), lowBound=0, cat="Integer")

ship_products = pulp.LpVariable.dicts("Ship_Products", (plants, distributors, products), lowBound=0, cat="Integer")
buy_materials = pulp.LpVariable.dicts("Buy_Materials", (suppliers, materials), lowBound=0, cat="Integer")
ship_materials = pulp.LpVariable.dicts("Ship_Materials", (suppliers, plants, materials), lowBound=0, cat="Integer")

for plant in plants:
    total_hours = pulp.lpSum([processing_time[plant, product] * make[plant][product] for product in products])
    model += total_hours <= plant_capacity[plant]
    
for distributor in distributors:
    for product in products:
        total_recieved = pulp.lpSum([ship_products[plant][distributor][product] for plant in plants])
        model += total_recieved <= market_demand[distributor, product]

for plant in plants:
    for product in products:
        total_shipped = pulp.lpSum([ship_products[plant][distributor][product] for distributor in distributors])
        model += total_shipped == make[plant][product]

for plant in plants:
    for material in materials:
        required_material = pulp.lpSum(to_tenths(material_needed[product, material]) * make[plant][product] for product in products)
        received_material = pulp.lpSum(ship_materials[supplier][plant][material] for supplier in suppliers)
        model += received_material == required_material

for supplier in suppliers:
    for material in materials:
        total_shipped = pulp.lpSum(ship_materials[supplier][plant][material] for plant in plants)
        model += buy_materials[supplier][material] == total_shipped
        model += buy_materials[supplier][material] <= to_tenths(
            material_limit[supplier, material]
        )

revenue = pulp.lpSum(
    selling_price[distributor, product] * ship_products[plant][distributor][product]
    for plant in plants
    for distributor in distributors
    for product in products
)

purchase_expense = pulp.lpSum(
    material_cost[supplier, material] * buy_materials[supplier][material] / tenth
    for supplier in suppliers
    for material in materials
)

raw_shipping_expense = pulp.lpSum(
    shipping_cost[material] * get_distance(supplier, plant) * ship_materials[supplier][plant][material] / tenth
    for supplier in suppliers
    for plant in plants
    for material in materials
)

labor_expense = pulp.lpSum(
    processing_time[plant, product] * make[plant][product] * hourly_wage[plant][product]
    for plant in plants
    for product in products
)

product_shipping_expense = pulp.lpSum(
    shipping_cost[material] * get_distance(plant, distributor) * ship_products[plant][distributor][product] 
    for plant in plants
    for distributor in distributors
    for product in products
)

total_expenses = labor_expense + product_shipping_expense + purchase_expense + raw_shipping_expense
profit = revenue - total_expenses
model += profit

solver = pulp.PULP_CBC_CMD(msg=False)
status = model.solve(solver)

if pulp.LpStatus[status] != "Optimal":
    raise ValueError("No optimal solution found. Status: {}".format(pulp.LpStatus[status]))

def get_value(variable):
    value = pulp.value(variable)

print("Profit:", round(get_value(profit), 2))

template = find_file("template")
workbook = load_workbook(template)
sheet_name = None

for name in workbook.sheetnames:
    if name.lower() == "solution":
        sheet_name = name
        break
worksheet = workbook[sheet_name]

def split_label(label, separators):
    for separator in separators:
        if separator in label:
            first, second = label.split(separator, 1)
            return first.strip(), second.strip()

for row in range(2, worksheet.max_row + 1):
    category = worksheet.cell(row, 1).value
    label = worksheet.cell(row, 2).value
    
    if category is None or label is None:
        continue
    category = str(category).strip()
    label = str(label).strip()
    quantity = 0
    
    if category == "supply":
        supplier, material = split_label(label, ["\\"])
        quantity = get_value(buy_materials[supplier][material]) / tenth
    elif category == "ship (supply->proc)":
        route, material = split_label(label, ["\\"])
        supplier, plant = split_label(route, ["→","->"])
        quantity = get_value(ship_materials[supplier][plant][material]) / tenth
    elif category == "process":
        plant, product = split_label(
            label, ["\\"]
        )
        quantity = get_value(make[plant][product])

    elif category == "ship (proc->distrib)":
        route, product = split_label(
            label, ["\\"]
        )
        plant, distributor = split_label(
            route, ["→", "->"]
        )
        quantity = get_value(
            ship_products[plant][distributor][product]
        )
    worksheet.cell(row, 3).value = round(quantity, 1)

result = output / "optimized_solution.xlsx"
workbook.save(result)