import urllib.request, json

ACCESS_KEY = "5754cb3f-0f1b-4d62-843d-df87a3aa8e55"
BASE_URL = "https://eservice.ura.gov.sg/uraDataService/invokeUraDS/v1"
TOKEN_URL = "https://eservice.ura.gov.sg/uraDataService/insertNewToken/v1"
RECENT = {"0326", "0426", "0526"}

def get_token():
    req = urllib.request.Request(TOKEN_URL, headers={"AccessKey": ACCESS_KEY})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())["Result"]

def fetch_batch(token, batch):
    url = f"{BASE_URL}?service=PMI_Resi_Transaction&batch={batch}"
    req = urllib.request.Request(url, headers={"AccessKey": ACCESS_KEY, "Token": token})
    with urllib.request.urlopen(req) as r:
        return json.loads(r.read())

token = "MdHr6eCXe522meJm468dG44ft8eatr8Rb2EcCamaVfZVbWKda1uUqRxby7gky5X48wVPyjfgK75wh5mR3FKJ85bcMdtmrddpQ7-2"
print(f"Token: {token[:20]}...")

data = fetch_batch(token, 1)
total = int(data.get("TotalBatch", 1))
print(f"Total batches: {total}")

all_results = list(data.get("Result", []))
for b in range(2, total + 1):
    d = fetch_batch(token, b)
    all_results.extend(d.get("Result", []))
    if b % 5 == 0:
        print(f"  Fetched {b}/{total} batches...")

d20 = []
for prop in all_results:
    for t in prop.get("transaction", []):
        if t.get("district") == "20" and t.get("contractDate") in RECENT:
            price = int(t.get("price", 0))
            area = float(t.get("area", 1))
            d20.append({
                "date": t.get("contractDate"),
                "project": prop.get("project", ""),
                "street": prop.get("street", ""),
                "type": t.get("propertyType", ""),
                "area_sqm": area,
                "price": price,
                "psf": round(price / area / 10.764),
                "sale_type": {"1":"New Sale","2":"Sub Sale","3":"Resale"}.get(t.get("typeOfSale"), ""),
                "tenure": t.get("tenure", ""),
                "floor": t.get("floorRange", "-"),
            })

d20.sort(key=lambda x: (x["date"], x["project"]), reverse=True)
print(f"\nD20 Transactions — Mar to May 2026 ({len(d20)} found)\n")
print(f"{'Date':<6} {'Project':<38} {'Type':<18} {'Area(sqm)':>9} {'Price':>12} {'PSF':>6} {'Floor':<8} {'Sale':<9} {'Tenure'}")
print("-"*125)
for t in d20:
    print(f"{t['date']:<6} {t['project'][:37]:<38} {t['type'][:17]:<18} {t['area_sqm']:>9.1f} {t['price']:>12,} {t['psf']:>6,} {t['floor']:<8} {t['sale_type']:<9} {t['tenure']}")
