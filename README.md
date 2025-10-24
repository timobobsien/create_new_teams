# Create new Teams with Powershell

This script creates new Teams via PowerShell in a specific manner.
We are an IT consulting company in Germany with an ISO27001 certification, so we have some criteria like permissions, standardization and data categorization.

This script loads all needed information from a CSV and creates in a loop all MS Teams. 
For our needs, we have teams like 

- Customer (Customer Documentation, Chats...)
- Internal (Teams for internal use like HR, Finance, NEST..)
- Private (For a single user)
- Project (For a single Project)
- Public (Shared with external users)

For each type, there are different channels/document folders/document tags (data categorization).
Our naming concept is like [company short name]-[type]-[purpose of use] like exe-internal-hr or exe-customer-xyzcompany

The CSV has the following formatting:

name;beschreibung;verfuegbarkeit;besitzer;mitglieder;ownerstrictlyconfidentialchannel;mitgliederstrictlyconfidentialchannel;
exe-internal-hr;example team hr;private;hruser1@example.com,hruser2@example.com;hruser3@example.com;hruser1@example.com;hruser2@example.com;

Feel free to use this script as an example and change it to your demands. 




