import pandas as pd
import ast

def process_excel(file_path):
    # Load the Excel file
    df = pd.read_excel(file_path, sheet_name='Sheet1')

    # Map TransparentLevel values
    transparency_mapping = {
        "No": 1,
        "Body": 2,
        "Full": 3
    }
    df['TransparentLevel'] = df['TransparentLevel'].map(transparency_mapping)

    # Process Colors column based on ColorListLength
    def process_colors(row):
        colors = ast.literal_eval(row['Colors'])  # Convert string to list
        if row['ColorListLength'] == 'RemoveSecond':
            if len(colors) > 2:  # If more than 2 colors, retain first and last
                colors = [colors[0], colors[-1]]
        elif row['ColorListLength'] == 'RemoveSecondThird':
            if len(colors) > 1:  # Ensure at least two elements exist
                colors.pop(1)  # Remove the second color (index 1)
            if len(colors) > 2:  # Ensure at least three elements exist for the third removal
                colors.pop(1)  # Remove the (new) third color (index 1 again)
        return colors

    df['Colors'] = df.apply(process_colors, axis=1)

    # Save the processed data to a new Excel file
    output_file = 'processed_switch_data.xlsx'
    df.to_excel(output_file, index=False)
    print(f"Processed data saved to {output_file}")

# Run the function with the provided file
process_excel('switch_data.xlsx')
