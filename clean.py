import pandas as pd, re
raw = pd.read_csv("titanic_raw.csv")
df = raw.copy()
log = []
def note(step, n, detail): log.append({"step": step, "rows_affected": int(n), "detail": detail})

# 1 duplicates
d = df.duplicated().sum(); df = df.drop_duplicates(); note("Remove exact duplicate rows", d, "Checked full-row duplicates")
dp = df["PassengerId"].duplicated().sum(); df = df.drop_duplicates("PassengerId"); note("Remove duplicate PassengerId", dp, "ID must be unique")
# 2 text trimming / casing
for c in ["Name","Sex","Ticket","Cabin","Embarked"]:
    df[c] = df[c].astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
df["Sex"] = df["Sex"].str.lower()
note("Trim spaces, normalise Sex to lowercase", len(df), "Name, Sex, Ticket, Cabin, Embarked")
# 3 split name
m = df["Name"].str.extract(r"^(?P<Surname>[^,]+),\s*(?P<Title>[^.]+)\.\s*(?P<GivenNames>.*)$")
df = pd.concat([df, m], axis=1)
df["Title"] = df["Title"].replace({"Mlle":"Miss","Ms":"Miss","Mme":"Mrs","Lady":"Rare","Countess":"Rare","the Countess":"Rare","Sir":"Rare","Don":"Rare","Dona":"Rare","Jonkheer":"Rare","Capt":"Rare","Col":"Rare","Major":"Rare","Rev":"Rare","Dr":"Rare"})
note("Split Name into Surname / Title / GivenNames, group rare titles", m["Title"].notna().sum(), "Mlle/Ms->Miss, Mme->Mrs, rare titles grouped")
# 4 missing values
miss_age = df["Age"].isna().sum()
df["Age_Missing"] = df["Age"].isna()
df["Age"] = df["Age"].fillna(df.groupby(["Title","Pclass"])["Age"].transform("median")).fillna(df["Age"].median()).round(0).astype(int)
note("Fill missing Age with median by Title and Class", miss_age, "Original missing rows flagged in Age_Missing")
miss_emb = df["Embarked"].isna().sum(); df["Embarked"] = df["Embarked"].fillna(df["Embarked"].mode()[0])
note("Fill missing Embarked with most common port", miss_emb, "Mode = S")
miss_cab = df["Cabin"].isna().sum(); df["Has_Cabin"] = df["Cabin"].notna(); df["Deck"] = df["Cabin"].str[0].fillna("Unknown")
note("Cabin: keep text, add Deck and Has_Cabin", miss_cab, "77% missing, so not imputed")
# 5 labels and types
df["Embarked"] = df["Embarked"].map({"S":"Southampton","C":"Cherbourg","Q":"Queenstown"})
df["Survived"] = df["Survived"].map({0:"No",1:"Yes"})
df["Class"] = df["Pclass"].map({1:"First",2:"Second",3:"Third"})
df["Fare"] = df["Fare"].round(2)
df["FamilySize"] = df["SibSp"] + df["Parch"] + 1
df["AgeGroup"] = pd.cut(df["Age"], [-1,12,19,39,59,120], labels=["Child","Teen","Adult","Middle-aged","Senior"])
note("Readable labels and new columns", len(df), "Embarked, Survived, Class names, FamilySize, AgeGroup")
# 6 outliers
q1,q3 = df["Fare"].quantile([.25,.75]); hi = q3+1.5*(q3-q1); df["Fare_Outlier"] = df["Fare"]>hi
note("Flag high Fare outliers (1.5 IQR), not removed", df["Fare_Outlier"].sum(), f"Threshold {hi:.2f}")
cols = ["PassengerId","Surname","Title","GivenNames","Sex","Age","Age_Missing","AgeGroup","Class","Survived","SibSp","Parch","FamilySize","Ticket","Fare","Fare_Outlier","Deck","Has_Cabin","Embarked"]
df = df[cols]
df.to_csv("titanic_clean.csv", index=False)
logdf = pd.DataFrame(log)
summ = pd.DataFrame({"Metric":["Rows in","Rows out","Columns in","Columns out","Missing cells before","Missing cells after"],
  "Value":[len(raw),len(df),raw.shape[1],df.shape[1],int(raw.isna().sum().sum()),int(df.isna().sum().sum())]})
before = raw.isna().sum().rename("Missing_before").to_frame()
with pd.ExcelWriter("titanic_cleaned.xlsx", engine="openpyxl") as w:
    summ.to_excel(w,"Summary",index=False); logdf.to_excel(w,"Cleaning_Log",index=False)
    before.to_excel(w,"Missing_Before"); df.to_excel(w,"Clean_Data",index=False)
    raw.head(50).to_excel(w,"Raw_Sample_50",index=False)
    for ws in w.book.worksheets:
        ws.freeze_panes="A2"
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(45, max(len(str(c.value or "")) for c in col[:60])+2)
print(summ.to_string(index=False)); print(logdf.to_string(index=False))
