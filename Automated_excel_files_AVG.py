import os
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns

# Set the backend for Matplotlib to work with Streamlit
plt.switch_backend("Agg")

def calculate_average(files_folder, column_name, row_index=None):
    files = os.listdir(files_folder)
    all_averages = []

    for file in files:
        if file.endswith(".xlsx"):
            file_path = os.path.join(files_folder, file)
            df = pd.read_excel(file_path)
            if column_name in df.columns:
                if row_index is None:
                    # Calculate column average for the entire column
                    column_average = df[column_name].mean()
                else:
                    # Calculate the average of the specific cell within the row if the index is valid
                    try:
                        cell_value = df.at[row_index, column_name]
                        column_average = cell_value
                    except KeyError:
                        column_average = None  # Handle invalid row index

                all_averages.append(column_average)

    return all_averages


def main():
    st.title("Excel Column Averages by Omar Adel Atito")

    folder_path = st.text_input("Enter the folder path:")
    column_to_average = st.text_input("Enter the column name to average:")
    specific_row_index = st.text_input("Enter the row index for cell average (leave empty for column average):")

    if st.button("Calculate Averages"):
        if os.path.exists(folder_path):
            if specific_row_index == "":
                specific_row_index = None
            else:
                specific_row_index = int(specific_row_index)

            averages = calculate_average(folder_path, column_to_average, specific_row_index)
            file_names = [file for file in os.listdir(folder_path) if file.endswith(".xlsx")]

            if len(file_names) == len(averages):
                new_file = pd.DataFrame({"File": file_names, "Average": averages})
                st.write(new_file)

                # Create a plot using seaborn
                plt.figure(figsize=(10, 6))
                sns.barplot(x="File", y="Average", data=new_file)
                plt.xlabel("File")
                plt.ylabel("Average")
                plt.title(f"Averages of '{column_to_average}' Column by Omar Adel Atito")
                plt.xticks(rotation=45)

                # Pass the figure object explicitly to st.pyplot()
                st.pyplot(plt.gcf())  # Get current figure

            else:
                st.error("Error: The number of averages and file names do not match.")
        else:
            st.error("Invalid folder path. Please enter a valid path.")

if __name__ == "__main__":
    main()
