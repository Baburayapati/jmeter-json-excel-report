
import json
import pandas as pd
import re

def is_transaction(name):
    return bool(re.match(r"^T\d{2}", str(name)))

def split_feature(name):
    return str(name).split("/")[0]

def load_data(path):
    with open(path) as f:
        data = json.load(f)
    rows = []
    for k,v in data.items():
        v['transaction'] = v.get('transaction', k)
        rows.append(v)
    return pd.DataFrame(rows)

def to_sec(df):
    df['avg_sec'] = df['meanResTime']/1000
    df['min_sec'] = df['minResTime']/1000
    df['max_sec'] = df['maxResTime']/1000
    return df

def bucket(val, askai):
    if askai:
        if val <=10: return "0-10"
        elif val <=20: return "10-20"
        elif val <=30: return "20-30"
        else: return ">30"
    else:
        if val <=2: return "0-2"
        elif val <=4: return "3-4"
        elif val <=6: return "4-6"
        else: return ">6"

def calc_track(df):
    df = df[~df['transaction'].apply(is_transaction)].copy()
    df = to_sec(df)
    df['Feature'] = df['transaction'].apply(split_feature)

    result = []

    for feature, g in df.groupby("Feature"):
        askai = str(feature).upper().startswith("ASKAI")
        total = len(g)

        for metric, col in [("Avg","avg_sec"),("Min","min_sec"),("Max","max_sec")]:
            counts = {"b1":0,"b2":0,"b3":0,"b4":0}
            for v in g[col]:
                b = bucket(v, askai)
                if b in ["0-2","0-10"]: counts["b1"]+=1
                elif b in ["3-4","10-20"]: counts["b2"]+=1
                elif b in ["4-6","20-30"]: counts["b3"]+=1
                else: counts["b4"]+=1

            result.append([
                feature, metric,
                round(counts["b1"]/total*100,2),
                round(counts["b2"]/total*100,2),
                round(counts["b3"]/total*100,2),
                round(counts["b4"]/total*100,2),
                round(g['max_sec'].max(),2)
            ])

    return pd.DataFrame(result, columns=[
        "Track","Metric","Bucket1 %","Bucket2 %","Bucket3 %","Bucket4 %","Max Seconds"
    ])

def generate(json_path, out):
    df = load_data(json_path)
    track = calc_track(df)
    with pd.ExcelWriter(out) as w:
        track.to_excel(w, sheet_name="Track_Comparison", index=False)

if __name__ == "__main__":
    import sys
    generate(sys.argv[1], sys.argv[2])
