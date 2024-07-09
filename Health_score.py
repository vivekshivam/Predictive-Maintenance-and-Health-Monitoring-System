import pandas as pd
from datetime import datetime, timedelta
from sklearn.linear_model import LinearRegression
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
import tkinter as tk
from tkinter import messagebox,ttk

# Step 1: Read all sheets from the Excel file into a single DataFrame
excel_file = 'classified_data_with_sheets.xlsx'

def load_all_data(excel_file):
    all_sheets = pd.read_excel(excel_file, sheet_name=None)
    df = pd.concat(all_sheets.values(), ignore_index=True)
    return df
# Step 2: Convert 'Notif.date' and 'Created at' columns to datetime format
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

# Step 3: Define the function to predict the next maintenance date and calculate health score
def predict_next_maintenance_date(equipment_id):
    df = load_all_data(excel_file)
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
        raise ValueError(f"No data found for the specified equipment ID '{equipment_id}'.")

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

    # Calculate health score based on average timedelta
    health_score = max(0, min(100, (avg_timedelta.days / 7) * 20))

    # Get previous maintenance events
    previous_events = "\n".join(
        f"Date: {row['Notif.date'].strftime('%Y-%m-%d')} | Category: {row['Category']}"
        for idx, row in equipment_data.iterrows()
    )
    return predicted_date, previous_events, health_score

# GUI Application
class MaintenancePredictorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Maintenance Predictor")
        self.geometry("600x400")

        self.label_equipment_id = tk.Label(self, text="Enter Equipment ID:")
        self.label_equipment_id.pack(pady=10)
        
        self.entry_equipment_id = tk.Entry(self)
        self.entry_equipment_id.pack(pady=10)

        self.button_predict = tk.Button(self, text="Predict Maintenance Date", command=self.predict_maintenance)
        self.button_predict.pack(pady=10)

        self.text_result = tk.Text(self, wrap='word', height=15, width=70)
        self.text_result.pack(pady=10)

    def predict_maintenance(self):
        equipment_id = self.entry_equipment_id.get().strip()
        if not equipment_id:
            messagebox.showerror("Input Error", "Please enter a valid Equipment ID.")
            return

        try:
            predicted_date, previous_events, health_score = predict_next_maintenance_date(equipment_id)
            if health_score < 10:
                health_status = "Immediate service required"
            elif health_score == 100:
                health_status = "Excellent"
            elif health_score > 50:
                health_status = "Good"
            
            else:
                health_status = "Service required"

            result = (f"Equipment ID: {equipment_id}\n"
                      f"Next Predicted Maintenance Date: {predicted_date}\n"
                      f"Health Score: {health_score} ({health_status})\n\n"
                      "Previous Maintenance Events:\n"
                      f"{previous_events}")
            self.text_result.delete(1.0, tk.END)
            self.text_result.insert(tk.END, result)
        except ValueError as e:
            messagebox.showerror("Data Error", str(e))

if __name__ == "__main__":
    app = MaintenancePredictorApp()
    app.mainloop()
