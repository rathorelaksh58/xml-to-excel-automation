import os
import xml.etree.ElementTree as ET
import pandas as pd

# 👉 Apna folder path yahan daal
folder_path = r"C:\Users\AGRA_ANA\Downloads\XML_4_2026_1777278684\XML_4_2026\ALL_METER"

data = []

for file in os.listdir(folder_path):
    # 👉 .xml / .XML dono handle karega
    if file.lower().endswith(".xml") and not file.startswith("~"):
        file_path = os.path.join(folder_path, file)
        
        try:
            tree = ET.parse(file_path)
            root = tree.getroot()

            g1 = root.find(".//G1")
            g2 = root.find(".//G2")
            g3 = root.find(".//G3")
            g4 = root.find(".//G4")

            data.append({
                "G1": g1.text.strip() if g1 is not None and g1.text else "",
                "G2": g2.text.strip() if g2 is not None and g2.text else "",
                "G3": g3.text.strip() if g3 is not None and g3.text else "",
                "G4": g4.text.strip() if g4 is not None and g4.text else ""
            })

        except Exception as e:
            print(f"❌ Error in file {file}: {e}")

# 👉 Excel file bana do
output_file = os.path.join(folder_path, "output.xlsx")

df = pd.DataFrame(data)
df.to_excel(output_file, index=False)

print("✅ Done bhai! Excel yahan bani hai:")
print(output_file)