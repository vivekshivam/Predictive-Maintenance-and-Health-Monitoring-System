import pandas as pd

# Read the Excel file into a DataFrame
df = pd.read_excel('Dataset.xlsx', sheet_name='Sheet1')

# Define the columns to drop
columns_to_drop = ['Notification', 'Notifictn type', 'Order', 'Breakdown', 
                   'Created By', 'Changed By', 'Changed On', 'Changed at', 
                   'Notif. Time', 'Plt for WorkCtr','Reported by', 'User status','Cost Center','Req. start']

# Drop the columns from the DataFrame
df = df.drop(columns=columns_to_drop, errors='ignore') 

#categorize
categories_keywords = {
    "Specialized Maintenance": [
        "steam trap", "thermography", "centrifuge", "silica gel", "fan motor",
        "lubrication system", "heat exchanger", "distillation column", 
        "pump overhaul", "valve maintenance", "compressor inspection"
    ],
    "Repair": [
        "leak", "faulty", "repair", "rectify","rectified", "rectification", "corroded",
        "weld repair", "valve repair", "pipeline repair", "pump repair", 
        "equipment failure","leak", "faulty", "repair", "rectify", "rectification", "remove", "seepage"
    ],
    "Replace": [
        "to be replaced", "replacement", "blinding", "damaged", "deblinding",
        "sleeving", "replace", "siding", "worn-out", "obsolete equipment","to be replaced", "replacement", "blinding", "damaged", "deblinding", "sleeving", "replace", "siding"
    ],
    "Preventive Maintenance": [
         "preventive", "maintenance", "calibration", "inspection",
        "pm", "bmr", "lubrication schedule", "equipment check-up", 
        "corrosion prevention","pm","pm of","pm job ","pm plan","maintenence job","maint job","cleaning", "preventive", "maintenance", "calibration", "inspection", "bmr"
    ],
    "General Maintenance": [
        "clean","cleaning", "top up", "maintenance", "clearing", "testing", "painting",
        "lubrication", "equipment adjustment", "filter replacement", 
        "routine check-up","clean", "top up", "maintenance","maint", "maint of", "clearing", "testing", "painting"
    ],
    "Testing and Inspection": [
        "checking", "check", "checked", "testing", "inspection", "thermography","checking", "check","monitoring","monitor", "checked", "testing", "inspection", "thermography",
        "calibration", "pressure test", "non-destructive testing",
        "integrity assessment","checking", "check","monitoring","monitor", "checked", "testing", "inspection", "thermography", "calibration"
    ],
    "Installation and Setup": [
        "connection", "disconnect", "setup", "installation", "commissioning",
        "decommissioning", "equipment installation", "new system setup",
        "facility setup", "piping installation","connection", "disconnect", "setup", "installation"
    ]
}

# Function to classify tasks based on description
def classify_task(description):
    description = str(description).lower() 
    
    for category, keywords in categories_keywords.items():
        for keyword in keywords:
            if keyword in description:
                return category
    
    return 'Other Classification'  
# Define a function to classify WorkCtr into branches based on keywords
def classify_work_center(work_center):
    keywords = {
        'electrical': ['elec'],
        'mechanical': ['mech'],
        'qc': ['qc'],
        'civil': ['civil'],
        'telecom': ['tele'],
        'f&s': ['fire', 'safety', 'f&s'],
        'inspection': ['insp'],
        'instrumentation': ['inst']
    }

    for branch, keywords_list in keywords.items():
        for keyword in keywords_list:
            if keyword in work_center.lower():
                return branch

    return None


# Apply classification function to DataFrame column
df['Category'] = df['Description'].apply(classify_task)

# Convert 'Notif.date' column to datetime format and format as dd-mm-yyyy
df['Notif.date'] = pd.to_datetime(df['Notif.date'], format='%Y%m%d').dt.strftime('%d-%m-%Y')

df.to_excel('classified_data.xlsx', index=False)

print("Classification and saving to Excel completed.")
