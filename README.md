# CoverLetter_Generator
## CoverLetter_Generator is a project based on gpt4o that autocompletes your cover letter
1. Modify your config:
    Go to "change_settings.ipynb" to get details.

2. Add job descriptions
    Add the job description into job_description.txt

3. Attach your resume
    Add your resume into Resume dir, then modify the resume parameter in generate_coverletter.ipynb

4. generate your cover letter:
    Go to "generate_coverletter.ipynb" to get details.

Remember to add your openai api_key into environment variables.

If your api_key in your environment variables cannot be found, add a backup at /my/envs/.env
Then the dotenv will solve this problem.
