import csv

with open('Upload-29-03-2023-20-02.csv', 'r', encoding='utf-8') as csv_file:
    print(type(csv_file))


files = {
    "csv_file": (
        "Upload-29-03-2023-20-02.csv.csv",
        open("Upload-29-03-2023-20-02.csv", "rb"),
        "text/csv",
        {"Expires": "0"},
    )
}