from openai import OpenAI
import json


client = OpenAI(
  api_key = 'sk-proj-vwLdoSbknx6deVWJ53e4T3BlbkFJbE7GObyEj4AHJNoqpKgM'
)

def process_mail_gpt(txt):

    prompt = """Parse the start date, end date and a list of clients from this email,
            return the response as json i can parse, make sure the clients are formatted as ['clients']
            start date is ['OD'] and end date is ['DO'] ALWAYS NAME THE START DATE ['OD'] AND END DATE ['DO']
            I AM PARSING THIS JSON IT DOESNT WORK IF U MESS UP THE KEYS.
            also make sure the dates are formated like dd.mm.yyyy.
            Capitalize the first and last name everything else should be lowercase, keep čđšžć"""
    message = f"{prompt} {txt}"


    completion = client.chat.completions.create(
      model="gpt-3.5-turbo",
      messages=[
        {"role": "user", "content": message}
      ]
    )

    #print(completion.choices[0].message.content)
    response = completion.choices[0].message.content
    return response
