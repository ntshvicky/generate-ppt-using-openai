# import the OpenAI Python library for calling the OpenAI API
from openai import OpenAI
import os

from dotenv import load_dotenv
load_dotenv()

client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY", os.getenv("OPENAPI_KEY")))


topic = "Future of Generative AI in Business"
num_slides = 12
content_prompt = f'''Generate content for a PowerPoint presentation on "{topic}". 
                
                The presentation should consist of {num_slides} slides with the following details:
                
                Preferred Structure:
                    Slide 1: Introduction(add small introduction on topic)
                    Slide 2-{num_slides-2}: Main Content (Different aspects of the impact)
                    Slide {num_slides-1}: Case Studies or Notable Examples
                    Slide {num_slides}: Conclusion
                
                Slide Details: For each slide in the PowerPoint presentation provide a concise bullet point followed by a brief explanation of 2 to 3 sentences. 
                Ensure that each bullet point is clear and informative, covering key aspects of {topic}. 
                The explanations should elaborate on the bullet point, providing context or additional details suitable for a professional audience. 
                Each slide should contain 5 to 6 of these bullet points with explanations, making the content informative and engaging.
                
                The total content should contain more than 1000 but less than 1200 tokens.
                    '''


# example with a system message
response = client.chat.completions.create(
    model="gpt-4",
    messages=[
        {"role": "user", "content": content_prompt},
    ],
    temperature=0,
)
#print(content_prompt)
print(response.choices[0].message.content)

with open("output.txt", "w") as f:
    f.write(response.choices[0].message.content)