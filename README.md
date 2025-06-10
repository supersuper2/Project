Setup Instructions

1. Install Python 3.8+

Ensure Python is installed:

`
python --version
`

2. Create a virtual environment
   
`
python -m venv venv
`

`
source venv/bin/activate  # or venv\Scripts\activate on Windows
`

3. Install requirements

`
pip install -r requirements.txt
`

4. Login to Azure CLI
   
`
az login
`

5. Replace Azure Subscription(Management) ID
`
subscription_id = "YOUR_ID"
`

6. Configure Email for Sending Reports:

`
msg["From"] = "YOUR_EMAIL@gmail.com"
`

`
server.login("YOUR_EMAIL@gmail.com", "YOUR_APP_PASSWORD")  # App password
`

7. Update email_mapping.json with real email addresses

8. Run program

