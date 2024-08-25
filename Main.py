import requests
import pandas as pd
from docx import Document


# Function to call the OpenAI ChatGPT API
def call_chatgpt(prompt, openai_api_key):
    url = "https://api.openai.com/v1/chat/completions"
    headers = {
        "Authorization": f"Bearer {openai_api_key}",
        "Content-Type": "application/json"
    }
    data = {
        "model": "gpt-4o-mini",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 16000  # Adjust as needed
    }
    response = requests.post(url, headers=headers, json=data)

    # Check for errors in the response
    if response.status_code == 200:
        response_json = response.json()
        if 'choices' in response_json and len(response_json['choices']) > 0:
            return response_json['choices'][0]['message']['content']
        else:
            return "Error: No choices found in the API response."
    else:
        return f"Error: API request failed with status code {response.status_code}. Response: {response.text}"


# Function to write results to a Word document
def write_results_to_docx(filename, data):
    doc = Document()
    doc.add_heading('ChatGPT Questions and Responses', level=1)

    # Loop through data (list of tuples) and write each question/response to the document
    for i, (prompt, response) in enumerate(data):
        doc.add_heading(f"Question {i + 1}: {prompt}", level=2)
        doc.add_paragraph(f"ChatGPT Response: {response}")
        #doc.add_paragraph("\n")  # Blank line between entries

    doc.save(filename)


# Main function to read prompts from Excel and store responses in a Word document
def main():
    openai_api_key = "Enter your API Code"  # Replace with your OpenAI API key
    input_excel_file = "prompts.xlsx"  # Replace with your Excel file containing prompts
    output_file = "ChatGPT_Responses.docx"

    # Read Excel file
    df = pd.read_excel(input_excel_file)

    # Assuming the prompts are in a column named 'Prompt'
    if 'Prompt' not in df.columns:
        print("Error: Excel file does not contain a 'Prompt' column.")
        return

    prompts = df['Prompt'].tolist()

    # Store results (prompt, response) in a list
    results = []

    # Make API calls for each prompt
    for prompt in prompts:
        print(f"Processing: {prompt}")
        response = call_chatgpt(prompt, openai_api_key)
        results.append((prompt, response))  # Store prompt and response

    # Write all results to a Word document
    write_results_to_docx(output_file, results)

    print(f"Responses saved in {output_file}")

if __name__ == "__main__":
    main()
