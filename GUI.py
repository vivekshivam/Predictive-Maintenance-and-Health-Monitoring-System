import pandas as pd
from datetime import datetime, timedelta
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
import tkinter as tk
from tkinter import ttk, messagebox

# Step 1: Read data from Excel file
excel_file = 'classified_data_with_sheets.xlsx'

# Step 2: Function to load the data for a given category
def load_data(category_name):
    category_safe_name = ''.join(c if c.isalnum() else '_' for c in category_name)
    df = pd.read_excel(excel_file, sheet_name=category_safe_name)
    return df

# Step 3: Convert 'Notif.date' and 'Created at' columns to datetime format
def preprocess_data(df):
    df['Notif.date'] = pd.to_datetime(df['Notif.date'], dayfirst=True)
    df.sort_values(by='Notif.date', inplace=True)
    df.reset_index(drop=True, inplace=True)
    df['Created at'] = pd.to_datetime(df['Notif.date'].dt.date.astype(str) + ' ' + df['Created at'])
    df['Month'] = df['Notif.date'].dt.month
    df['Hour'] = df['Created at'].dt.hour
    df['Minute'] = df['Created at'].dt.minute
    df['Second'] = df['Created at'].dt.second
    return df

# Step 4: Define the function to predict the next maintenance date
def predict_next_maintenance_date(equipment_id, category_name):
    df = load_data(category_name)
    df = preprocess_data(df)

    # Define features and target variable
    X = df[['Month', 'Hour', 'Minute', 'Second', 'System status', 'Category', 'Branch']]
    y = df['Notif.date']

    # Define numerical and categorical columns
    categorical_cols = ['System status', 'Category', 'Branch']
    numerical_cols = ['Month', 'Hour', 'Minute', 'Second']

    # Column transformer for preprocessing
    preprocessor = ColumnTransformer(
        transformers=[
            ('num', StandardScaler(), numerical_cols),
            ('cat', OneHotEncoder(handle_unknown='ignore'), categorical_cols)
        ])

    # Pipeline with preprocessing and model
    pipeline = Pipeline(steps=[
        ('preprocessor', preprocessor),
        ('regressor', LinearRegression())
    ])

    # Train the model
    pipeline.fit(X, y)

    # Filter data for the specified equipment ID
    equipment_data = df[df['Functional Loc.'] == equipment_id].copy()

    if equipment_data.empty:
        raise ValueError(f"No data found for the specified equipment ID '{equipment_id}' in category '{category_name}'.")

    # Get the most recent maintenance date and calculate average timedelta
    notification_dates = equipment_data['Notif.date']
    avg_timedelta = notification_dates.diff().mean()

    # Prepare input data for prediction
    last_maintenance_date = notification_dates.max()
    new_data = {
        'Month': [last_maintenance_date.month],
        'Hour': [equipment_data['Hour'].iloc[-1]],
        'Minute': [equipment_data['Minute'].iloc[-1]],
        'Second': [equipment_data['Second'].iloc[-1]],
        'System status': [equipment_data['System status'].iloc[-1]],
        'Category': [equipment_data['Category'].iloc[-1]],
        'Branch': [equipment_data['Branch'].iloc[-1]]
    }

    # Convert new_data to DataFrame
    new_df = pd.DataFrame(new_data)

    # Preprocess and predict
    predicted_timestamp = pipeline.predict(new_df)[0]

    # Convert predicted timestamp to datetime
    try:
        predicted_date = datetime.fromtimestamp(predicted_timestamp)
    except OSError:
        predicted_date = last_maintenance_date + avg_timedelta

    # Get previous maintenance events
    previous_events = "\n".join(
        f"Date: {row['Notif.date'].strftime('%Y-%m-%d')} | Category: {row['Category']}"
        for idx, row in equipment_data.iterrows()
    )

    return predicted_date, previous_events

# GUI Application
class MaintenanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Maintenance Prediction")

        # Create and place widgets
        ttk.Label(root, text="Equipment ID:").grid(column=0, row=0, padx=10, pady=10)
        self.equipment_id_entry = ttk.Entry(root)
        self.equipment_id_entry.grid(column=1, row=0, padx=10, pady=10)

        ttk.Label(root, text="Category:").grid(column=0, row=1, padx=10, pady=10)
        self.category_combobox = ttk.Combobox(root, values=self.get_categories())
        self.category_combobox.grid(column=1, row=1, padx=10, pady=10)

        self.predict_button = ttk.Button(root, text="Predict", command=self.predict)
        self.predict_button.grid(column=0, row=2, columnspan=2, pady=20)

        self.result_text = tk.Text(root, height=10, width=50)
        self.result_text.grid(column=0, row=3, columnspan=2, pady=10, padx=10)

    def get_categories(self):
        df = pd.read_excel(excel_file, sheet_name=None)
        return list(df.keys())

    def predict(self):
        equipment_id = self.equipment_id_entry.get()
        category_name = self.category_combobox.get()
        try:
            predicted_date, previous_events = predict_next_maintenance_date(equipment_id, category_name)
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"Next Predicted Maintenance Date: {predicted_date}\n\n")
            self.result_text.insert(tk.END, "Previous Maintenance Events:\n")
            self.result_text.insert(tk.END, previous_events)
        except ValueError as e:
            messagebox.showerror("Error", str(e))

# Run the GUI application
if __name__ == "__main__":
    root = tk.Tk()
    app = MaintenanceApp(root)
    root.mainloop()
