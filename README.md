# AdvisoryBoardEmailer scrubs spreadsheet for information and writes an email to gmail

### My use is to send the status of our company's Advisory Board jobs programaticaly and automaticaly. 

If you copy this repository, you will need to generate credentials from google's developer console and add to the project as client_secret.json.  When you run the code with a client_secret.json present, you will be redirected to a gmail signin page, after signing in, a .credentials folder will be generated with your oauth key.  After that, the emailer will work.