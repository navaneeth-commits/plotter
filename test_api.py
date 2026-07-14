import os
import sys
from google import genai

api_key = os.environ.get("GEMINI_API_KEY")
if not api_key:
    sys.exit("Set GEMINI_API_KEY in your environment before running this script.")

client = genai.Client(api_key=api_key)

try:
    response = client.models.generate_content(
        model='gemini-2.5-flash',
        contents='''You are a local data analyst.
Return STRICT JSON only.
Format:
{
  "insights": ["..."],
  "charts": [
    {
      "type": "bar",
      "x": "column",
      "y": "column",
      "group": null
    }
  ]
}
Dataset:
{"columns":[{"name":"test","type":"numeric"},{"name":"cat","type":"category"}], "sample_rows":[{"test":1,"cat":"A"}]}'''
    )
    print("OUTPUT:", response.text)
except Exception as e:
    print("ERROR:", str(e))
