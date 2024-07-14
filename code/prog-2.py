import google.generativeai as genai

# Configure your API key
api_key = "AIzaSyBIj_b-axuFTRVrc4ya65CCfazNejT6u-Y"
genai.configure(api_key=api_key)

# Load the generative model
model = genai.GenerativeModel(model_name="gemini-pro")

# Generate content
response = model.generate_content([data])

print(response.text)
