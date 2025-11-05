import pandas as pd
from openpyxl import load_workbook
import numpy as np
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score

#I. DATA REPROCESSING
    #1. Data Integration
file_path = r"C:/Users/levuo/PycharmProjects/PythonProject/MoMo_DA.xlsx"

        # Reading data from sheets
transactions_df = pd.read_excel(file_path, sheet_name='Data Transactions')
commission_df = pd.read_excel(file_path, sheet_name='Data Commission')
user_info_df = pd.read_excel(file_path, sheet_name='Data User_Info')

        # Merging data
merged_df = transactions_df.merge(commission_df, on='Merchant_id', how='left')
final_merged_df = merged_df.merge(user_info_df, on='User_id', how='left')

        # Load current workbook and add a new sheet
book = load_workbook(file_path)
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        # Writing merged data into the new sheet
    final_merged_df.to_excel(writer, sheet_name='Merged_Data', index=False)

    #2. Data cleaning
df = pd.read_excel(file_path, sheet_name='Merged_Data', engine='openpyxl')

        #Processing missing data
            # Handling blank cells in Purchase_status column
df['Purchase_status'].fillna('Not "Mua hộ"', inplace=True)

        #Processing outliers
            # Handling the cell '18_to_22'
mask_18_22 = df['Age'] == '18_to_22'
half_18_22 = mask_18_22.sum() // 2
df.loc[mask_18_22, 'Age'] = ([20] * half_18_22) + ([20] * (mask_18_22.sum() - half_18_22))

            # Handling the cell '23_to_27'
mask_23_27 = df['Age'] == '23_to_27'
half_23_27 = mask_23_27.sum() // 2
df.loc[mask_23_27, 'Age'] = ([25] * half_23_27) + ([25] * (mask_23_27.sum() - half_23_27))

            # Handling the cell '>37'
mask_over_37 = df['Age'] == '>37'
df.loc[mask_over_37, 'Age'] = np.random.randint(37, 71, size=mask_over_37.sum())

            # Handling the cell 'unknown'
mask_unknown = df['Age'] == 'unknown'
third_unknown = mask_unknown.sum() // 3
df.loc[mask_unknown, 'Age'] = (
    [20.5] * third_unknown +
    [24.5] * third_unknown +
    list(np.random.randint(37, 81, size=(mask_unknown.sum() - 2 * third_unknown)))
)

            # Handling the cell'28_to_32'
mask_28_32 = df['Age'] == '28_to_32'
quarter_28_32 = mask_28_32.sum() // 4
remaining = mask_28_32.sum() - 3 * quarter_28_32
df.loc[mask_28_32, 'Age'] = (
    [28] * quarter_28_32 +
    [32] * quarter_28_32 +
    [30] * quarter_28_32 +
    [30] * remaining
)

            # Handling the cell '33_to_37'
mask_33_37 = df['Age'] == '33_to_37'
quarter_33_37 = mask_33_37.sum() // 4
remaining = mask_33_37.sum() - 3 * quarter_33_37
df.loc[mask_33_37, 'Age'] = (
    [33] * quarter_33_37 +
    [37] * quarter_33_37 +
    [35] * quarter_33_37 +
    [35] * remaining
)
        # Processing inconsistent data
            # Normalizing Gender column data
df['Gender'] = df['Gender'].replace({
    'f': 'FEMALE',
    'female': 'FEMALE',
    'male': 'MALE',
    'Nữ': 'FEMALE',
    'Nam': 'MALE',
    'M': 'MALE'
})

            # Initializing list of provinces and cities in Vietnam (from Hanoi and Ho Chi Minh City)
vn_cities = [
    'An Giang', 'Ba Ria - Vung Tau', 'Bac Lieu', 'Bac Giang', 'Bac Kan', 'Bac Ninh', 'Ben Tre', 'Binh Dinh',
    'Binh Duong', 'Binh Phuoc', 'Binh Thuan', 'Ca Mau', 'Can Tho', 'Cao Bang', 'Da Nang', 'Dak Lak', 'Dak Nong',
    'Dien Bien', 'Dong Nai', 'Dong Thap', 'Gia Lai', 'Ha Giang', 'Ha Nam', 'Ha Tinh', 'Hai Duong', 'Hai Phong',
    'Hau Giang', 'Hoa Binh', 'Hung Yen', 'Khanh Hoa', 'Kien Giang', 'Kon Tum', 'Lai Chau', 'Lam Dong', 'Lang Son',
    'Lao Cai', 'Long An', 'Nam Dinh', 'Nghe An', 'Ninh Binh', 'Ninh Thuan', 'Phu Tho', 'Phu Yen', 'Quang Binh',
    'Quang Nam', 'Quang Ngai', 'Quang Ninh', 'Quang Tri', 'Soc Trang', 'Son La', 'Tay Ninh', 'Thai Binh',
    'Thai Nguyen', 'Thanh Hoa', 'Thua Thien Hue', 'Tien Giang', 'Tra Vinh', 'Tuyen Quang', 'Vinh Long', 'Vinh Phuc',
    'Yen Bai'
]

            # Normalizing Location column data
df['Location'] = df['Location'].replace({
    'HN': 'Ha Noi',
    'HCMC': 'Ho Chi Minh City'
})

                # Handling cells with values "Other Cities", "Other", "Unknown"
mask_replace_cities = df['Location'].str.strip().isin(['Other Cities', 'Other', 'Unknown'])
df.loc[mask_replace_cities, 'Location'] = np.random.choice(vn_cities, size=mask_replace_cities.sum())
#II. CARRYING OUT THE REQUIREMENTS IN THE EXCEL FILE
    #Part A:
                 # Adding Revenue column
                     #  Converting Amount column to numeric type (remove comma)
df['Amount'] = df['Amount'].replace({',': ''}, regex=True).astype(float)
                    # Calculating the Revenue column
df['Revenue'] = df['Amount'] * (df['Rate_pct'] / 100)

                     # Adding new column "Day of Week"
df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
df['Day of Week'] = df['Date'].dt.day_name()


                    # Extract the month and year for comparison
df['First_tran_date'] = pd.to_datetime(df['First_tran_date'], errors='coerce')
df['FirstYearMonth'] = df['First_tran_date'].dt.to_period('M').astype(str)


                # Create 'Type_user' column based on the given condition
                    # Creating YearMonth column from Date column for comparison
df['Month'] = df['Date'].dt.to_period('M').astype(str)
df['Type_user'] = df.apply(lambda row: 'New' if row['Month'] == row['FirstYearMonth'] else 'Current', axis=1)

                    # Drop helper columns
df.drop(columns=['Month', 'FirstYearMonth'], inplace=True)

                    # Writing the processed data to the 'Merged_Data' sheet itself
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Merged_Data', index=False)

                #Creating statistical tables
                    #The table shows the number of distinct users and the number of transactions over the months.
                        # Converting date column to datetime format
df['date'] = pd.to_datetime(df['Date'])

                        # Creating column 'month' in format 'YYYY-MM'
df['Months'] = df['date'].dt.strftime('%Y-%m')

                        # Calculating the number of different user_id and count the number of order_id by month
stats = df.groupby('month').agg(
    distinct_user_id=('User_id', 'The quantity of Distinct Users'),
    total_order_id=('order_id', 'The Quantity of Transactions')
).reset_index()

                        # Writing the results to the "Statistical table" sheet in the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    stats.to_excel(writer, sheet_name='Table 1', index=False)

                    #The table shows revenue of Merchant_names over the months
                        # Creating a pivot table: each row is Merchant_name, columns are months, value total revenue from column 'total_amount'
pivot_revenue = df.pivot_table(
    index='Merchant_name',
    columns='Months',
    values='Revenue',
    aggfunc='sum',
    fill_value=0
).reset_index()

                        # Writing the results to the "TABLE" sheet in the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    pivot_revenue.to_excel(writer, sheet_name='Table 2', index=False)

                     #The table shows revenue of places over the months
                        # Classifying cities into 3 groups: "Ha Noi", "Ho Chi Minh City" and "Other Cities"
def classify_city(city):
    if city == "Ha Noi":
        return "Ha Noi"
    elif city == "Ho Chi Minh City":
        return "Ho Chi Minh City"
    else:
        return "Other Cities"

df['city_group'] = df['Location'].apply(classify_city)

                            # Creating a pivot table: each row is a group of cities, the columns are the months, the values the total revenue
pivot_revenue = df.pivot_table(
    index='city_group',
    columns='Months',
    values='Revenue',
    aggfunc='sum',
    fill_value=0
).reset_index()

                            # Writing results in the "TABLE" sheet in the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    pivot_revenue.to_excel(writer, sheet_name='Table 3', index=False)
                    #The table of shows the quantity of new and curent users over the months
                            # Filtering rows with the value 'New' or 'Current' in the 'status' column
df_status = df[df['Type_user'].isin(['New', 'Current'])]

                            # Creating a statistics table: each row is a month, the columns 'New' and 'Current' are the corresponding number of rows
stats = df_status.groupby(['Months', 'Type_user']).size().unstack(fill_value=0).reset_index()

                            # If necessary, rename the columns for clarity (e.g. New -> New row number, Current -> Current row number)
stats = stats.rename(columns={'New': 'The Quantity of New Users', 'Current': 'The Quantity of Current Users'})

                            # Writing the results to the "SHOT" sheet in the Excel file
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    stats.to_excel(writer, sheet_name='Table 4', index=False)

                # Drawing charts in order to visualize data
                        #The line chart showing revenue over months
                            # Grouping data by Month and calculate total Revenue
grouped = df.groupby('Month').agg({'Revenue': 'sum'}).reset_index()

                            # Converting the Month column to a string type so matplotlib can handle it
grouped['Month'] = grouped['Month'].astype(str)

                            # Drawing a line graph
plt.figure(figsize=(10,6))
plt.plot(grouped['Month'], grouped['Revenue'], marker='o', color='blue')
plt.xlabel('Month')
plt.ylabel('Revenue')
plt.title('Revenue over months ')
plt.ticklabel_format(style='plain', axis='y')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

                        # The line chart showing average revenue by day of the week
df['Day_of_Week'] = df['date'].dt.day_name()

                            # Grouping by Day_of_Week and average revenue
avg_revenue = df.groupby('Day_of_Week')['Revenue'].mean().reset_index()

                            #Sorting by standard weekly order: Monday -> Sunday
days_order = ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday']
avg_revenue['Day_of_Week'] = pd.Categorical(avg_revenue['Day_of_Week'],
                                            categories=days_order,
                                            ordered=True)

avg_revenue = avg_revenue.sort_values('Day_of_Week')

                            #Drawing line graph with marker
plt.figure(figsize=(10,6))
plt.plot(avg_revenue['Day_of_Week'], avg_revenue['Revenue'], marker='o', color='blue',linestyle='-')
plt.xlabel('Day of Week')
plt.ylabel('Average Revenue')
plt.title('Average Revenue by Day of Week ')
plt.grid(True)

                    #The Column and line chart showing the number of distinct users and the number of transactions over months
                            # Removing rows without valid Date values (avoid NaN/NaT)
df = df.dropna(subset=['Date'])

                            # Grouping data by month and count values of 'User Id' and 'Order_id'
grouped = df.groupby('Month').agg({
    'User_id': 'count',
    'order_id': 'count'
}).reset_index()

                            # Renaming columns for ease of understanding
grouped.rename(columns={'User_id': 'Count_User', 'order_id': 'Count_Order'}, inplace=True)

                            # Filtering out NaT Month values (if any)
grouped = grouped[grouped['Month'] != 'NaT']

                            # Drawing a chart
fig, ax1 = plt.subplots(figsize=(10, 6))

                            # Drawing a column chart for the number of User_id
ax1.bar(grouped['Month'], grouped['Count_User'], color='skyblue', label='Number of User_ID')
ax1.set_xlabel('Months')
ax1.set_ylabel('Number of Distinct Users', color='blue')
ax1.tick_params(axis='y', labelcolor='blue')
ax1.set_xticklabels(grouped['Month'], rotation=45)

                            # Using the second axis to plot a line for the Order_id quantity
ax2 = ax1.twinx()
ax2.plot(grouped['Month'], grouped['Count_Order'], color='red', marker='o', label='Number of Order_ID')
ax2.set_ylabel('Number of Transactions', color='red')
ax2.tick_params(axis='y', labelcolor='red')

                            # Adding title and caption
plt.title('Number of Distinct Users and Transactions over months')
fig.tight_layout()
fig.legend(loc='upper left', bbox_to_anchor=(0.1, 0.9))
plt.show()

                    # The Combined column chart shows the number of new and current users over months
                        # Grouping data by Month and Type User, count the quantity
grouped = df.groupby(['Month', 'Type_user']).size().reset_index(name='Count')

                        # Pivot so that each Type_user value becomes a column, the value is Count
pivoted = grouped.pivot(index='Month', columns='Type_user', values='Count').fillna(0)

                        # Plotting a combined column chart (each column represents a Type_user in the same month)
pivoted.plot(kind='bar', figsize=(10, 6))

                        # Customizing axes, titles, legends
plt.xlabel('Month')
plt.ylabel('The quantity of new users and current users (persons)')
plt.title('The quantity of new users and current users over the months')
plt.legend(title='Type_user')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

                    #The Stacked column chart showing the number of transactions of Merchant_names over the months
                        # Grouping data by Month and Merchant_name, count the number of Order_id
grouped = df.groupby(['Month', 'Merchant_name'])['Order_id'].count().reset_index(name='Count_Order_id')

                        #  Pivot so that each Merchant Name is 1 column, each Month is 1 row
pivoted = grouped.pivot(index='Month', columns='Merchant_name', values='Count_Order_id').fillna(0)

                        # Drawing a stacked column chart
pivoted.plot(kind='bar', stacked=True, figsize=(18, 6))

                        # Customizing axis labels and titles
plt.xlabel('Month')
plt.ylabel('Total quantity of transactions')
plt.title('The number of transactions of Merchant_names over months')

                        # Showing caption
plt.legend(ncol=2, bbox_to_anchor=(1.05, 1), loc='upper left')

                        # Rotating X-axis label for readability
plt.xticks(rotation=45)

                        # Showing chart
plt.tight_layout()
plt.show()

                #The Combined line chart showing the number of transactions by Merchant_names over the months
                    # Grouping data by Month and Merchant Name, count the number of occurrences of each Merchant Name
grouped = df.groupby(['Month', 'Merchant_name']).size().reset_index(name='Count')

                    # Pivot data so that each row is 1 month, each column is 1 Merchant_name
pivoted = grouped.pivot(index='Month', columns='Merchant_name', values='Count').fillna(0)

bottom_merchant = pivoted.sum().idxmin()

                        # Drawing a combined line chart: each Merchant_name is a line
plt.figure(figsize=(10, 6))
for merchant in pivoted.columns:
    plt.plot(range(len(pivoted.index)), pivoted[merchant], marker='o', label=merchant)
    for i, value in enumerate(pivoted[merchant]):
        month_label = pivoted.index[i]
        if (merchant == 'Viettel') and month_label.endswith('-12'):
            offset = (0, -15)
            va = 'top'
        elif merchant == bottom_merchant:
            offset = (0, 4)
            va = 'bottom'
        else:
            offset = (0, 10)
            va = 'bottom'
        plt.annotate(
            f'{value:.0f}',
            xy=(i, value),
            xytext=offset,
            textcoords='offset points',
            ha='center',
            va=va,
            fontsize=9,
            color='black'
        )
plt.xlabel('Months')
plt.ylabel('The number of transactions')
plt.title('The number of Merchant_names transactions over months')
plt.xticks(range(len(pivoted.index)), pivoted.index, rotation=45)
plt.legend(title='Merchant_name', loc='upper left', bbox_to_anchor=(1.05, 1))
plt.tight_layout()
plt.show()


            #The Stacked area chart showing percentage of transactions by locality over months
                # Supposing the data has a column called "City". If the column has a different name, replace "City" with the correct name.
df['City_Category'] = df['Location'].apply(lambda x: x if x in ['Ha Noi', 'Ho Chi Minh City'] else 'Other Cities')

                # Grouping data by Month and City_Category, count number of transactions
grouped = df.groupby(['Month', 'City_Category']).size().reset_index(name='Count')

                # Calculating total transactions per month to calculate percentage
total_per_month = grouped.groupby('Month')['Count'].sum().reset_index(name='Total')
merged = pd.merge(grouped, total_per_month, on='Month')
merged['Percentage'] = merged['Count'] / merged['Total'] * 100

                # Pivot data: each row is 1 Month, each column is 1 City_Category with value as % transaction
pivot = merged.pivot(index='Month', columns='City_Category', values='Percentage').fillna(0)

                # Determining the full month range from January to December (here assume 2020)
full_index = pd.date_range(start="2020-01-01", end="2020-12-01", freq='MS').strftime('%Y-%m')

                # Reindex pivot to add missing months, fill in 0 if data is not available
pivot = pivot.reindex(full_index, fill_value=0)

                # Drawing a stacked area chart
plt.figure(figsize=(10, 6))
pivot.plot.area(alpha=0.8)

plt.title('Percentage of localities over months')
plt.xlabel('Months')
plt.ylabel('Percentage_rate (%)')
plt.legend(title='Places', loc='upper left', bbox_to_anchor=(1.05, 1))
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()


            # The Combined horizontal bar chart showing total revenue of localities over months
                # Grouping data by Month and City_Category, calculate total Revenue for each group
grouped = df.groupby(['Month', 'City_Category'])['Revenue'].sum().reset_index()

                # Pivot data: each row is 1 Month, each column is 1 City_Category with value is total Revenue
pivot = grouped.pivot(index='Month', columns='City_Category', values='Revenue').fillna(0)

                # If you want to ensure full 12 months are displayed (for example from 2020-01 to 2020-12), we reindex:
full_index = pd.date_range(start="2020-01-01", end="2020-12-01", freq='MS').strftime('%Y-%m')
pivoted = pivot.reindex(full_index, fill_value=0)

                # Drawing a combined horizontal bar chart (each row: 1 month, each bar: 1 group)
plt.figure(figsize=(12, 6))
pivoted.plot(kind='barh', figsize=(12, 6))
plt.xlabel('Total Revenue')
plt.ylabel('Months')
plt.title('Total Revenue of places over months')
plt.ticklabel_format(style='plain', axis='x')
plt.legend(title='Places', loc='lower right', bbox_to_anchor=(0,1))
plt.tight_layout()
plt.show()

            #The Pie chart showing percentage of types of user in the highest revenue month - December
                # Grouping data by Month and calculate total revenue
grouped = df.groupby('Month')['Revenue'].sum().reset_index()

                #Determining the month with the highest revenue
max_month = grouped.loc[grouped['Revenue'].idxmax(), 'Month']

                #Getting data for the month with the highest revenue
df_max = df[df['Month'] == max_month]

                #Grouping by Gender and Count
gender_counts = df_max['Gender'].vallue_counts()

                #Drawing a pie chart showing percentages
plt.figure(figsize=(8,8))
plt.pie(gender_counts, labels=gender_counts.index, autopct='%1.1%%', startangle=90)
plt.title('Percentage of Gender values for the month owning the highest revenue - December')
plt.tight_layout()
plt.show()


            #The Horizontal bar chart showing total revenue of Merchant_names in the month with the highest revenue
                # Grouping by Merchant_name to calculate total revenue for that month
merchant_revenue = df_max.groupby('Merchant_name')['Revenue'].sum().reset_index()

                # Sorting by revenue descending
merchant_revenue = merchant_revenue.sort_values(by='Revenue', ascending=True)

# Drawing a horizontal bar chart
plt.figure(figsize=(10, 6))
plt.barh(merchant_revenue['Merchant_name'], merchant_revenue['Revenue'], color='skyblue')
plt.xlabel('Revenue')
plt.ylabel('Merchant_name')
plt.title('Revenue of Merchant_names in the month owning the highest revenue - December')
plt.tight_layout()
plt.show()



#III. Appling Linear Regression Model to train and predict on Revenue dataset with MAE, MSE, RMSE and R-squared indexes calculated to evaluate model performance.

# Eliminating rows with missing revenue values
data = df.dropna(subset=['Revenue'])

# Separating X and y
X = data.drop(columns=["Revenue"])
y = data["Revenue"]



# Handling datetime columns in X
datetime_cols = X.select_dtypes(include=['datetime64']).columns

for col in datetime_cols:
    # Converting the column to datetime if not already, then convert to timestamp (number of seconds)
    X[col] = pd.to_datetime(X[col], errors='coerce')
    X[col] = X[col].apply(lambda x: x.timestamp() if pd.notnull(x) else np.nan)
    # Filling in missing values if any
    X[col].fillna(X[col].mean(), inplace=True)

# If there are categorical columns, convert them to dummy variables
X = pd.get_dummies(X)


# Splitting data into training set and testing set
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)

# Training Linear Regression Model
model = LinearRegression()
model.fit(X_train, y_train)

#  Predicting and evaluating the model on the test set
y_pred = model.predict(X_test)

mae = mean_absolute_error(y_test, y_pred)
mse = mean_squared_error(y_test, y_pred)
rmse = np.sqrt(mse)
r2 = r2_score(y_test, y_pred)

print("MAE:", mae)
print("MSE:", mse)
print("RMSE:", rmse)
print("R-squared:", r2)

