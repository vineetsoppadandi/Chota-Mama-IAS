import requests
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
        "max_tokens": 16384
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
def write_results_to_docx(filename, prompt, chatgpt_response):
    doc = Document()
    doc.add_heading(prompt, level=1)
    doc.add_paragraph(f"Question: {prompt}")
    doc.add_paragraph(f"ChatGPT Response: {chatgpt_response}")
    doc.save(filename)

# Main function
def main():
    openai_api_key = "sk-OXNsYNxSDx8kdtQxs9Xn7DrHwvdoYKx2MM8NGmHJp4T3BlbkFJ6X8wkotlK286QP4TUVmaxolwS0suY1gscjajNrWE4A-T"
    prompt = "In the context of Telangana india explain Information and Communication Technology (ICT) Policy of Telangana in alot of detail with links to learn more"
    output_file = "ChatGPT_Response.docx"

    # Call ChatGPT API
    response = call_chatgpt(prompt, openai_api_key)

    # Write question and response to Word document
    write_results_to_docx(output_file, prompt, response)

    print(f"Response saved in {output_file}")

if __name__ == "__main__":
    main()
