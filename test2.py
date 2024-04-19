import os
import openai

from dotenv import load_dotenv
load_dotenv()

# Initialize your OpenAI API key
openai.api_key = os.getenv("OPENAPI_KEY")

def generate_slide_titles(topic, num_slides):
    prompt = f"Create an outline with titles for a {num_slides}-slide PowerPoint presentation on the topic: '{topic}'."
    response = openai.Completion.create(
        engine="gpt-3.5-turbo-instruct",
        prompt=prompt,
        max_tokens=50 * num_slides  # Assuming each title would be around 50 tokens
    )
    titles = response.choices[0].text.strip().split('\n')
    return titles

def generate_slide_content(title):
    prompt = f"Generate a detailed content for a PowerPoint slide titled '{title}'. Provide a series of key points that expand on the topic."
    response = openai.Completion.create(
        engine="gpt-3.5-turbo-instruct",
        prompt=prompt,
        max_tokens=150  # Adjust the number of tokens as needed for detailed content
    )
    content = response.choices[0].text.strip()
    return content

def main():
    topic = input("Enter the topic for the PowerPoint presentation: ")
    num_slides = int(input("How many slides do you want to generate? "))
    
    print(f"Generating an outline for {num_slides} slides on the topic '{topic}'...")
    slide_titles = generate_slide_titles(topic, num_slides)
    
    for i, title in enumerate(slide_titles, start=1):
        print(f"Slide {i}: {title}")

    generate_content = input("Do you want to generate detailed content for each slide? (Y/N) ").strip().lower()
    
    if generate_content == 'y':
        for i, title in enumerate(slide_titles, start=1):
            print(f"\nGenerating content for Slide {i}: {title}...")
            content = generate_slide_content(title)
            print(f"Content for Slide {i}:\n{content}")

if __name__ == "__main__":
    main()
