# SLA_Calculations
parse excel spreadsheet to create SLA reports for Entech 

## How to Use 
1. Create Github Account, click Settings -> '<>'Developer Settings -> Personal Access Tokens -> Tokens (classic) -> Generate New Token (classic) -> Click "No Expiration" -> Generate Token
2. Copy token and save somewhere on machine (recommend emailing yourself or putting in a google doc/word doc)
3. Download Eclipse from https://www.eclipse.org/downloads/
4. Download Git from https://www.git-scm.com/downloads
5. Download Java from https://www.java.com/download/ie_manual.jsp or https://www.oracle.com/java/technologies/downloads/ (oracle.com is my recommendation)
6. Open Eclipse and go to the Git perspective. You can do this by clicking on Window -> Perspective -> Open Perspective -> Others..., then select Git and click OK.
7. In the Git perspective, click on Clone a Git repository in the Git Repositories view. If you don't see this view, you can open it by going to Window -> Show View -> Other..., then select Git -> Git Repositories and click OK.
8. In the Clone Git Repository dialog that appears, enter the URI of the repository (found clicking `<Code>` within repo).
9. Click Next to select the branches you want to clone, and then click Next again to set the local destination for the repository (want to clone main branch).
10. May need to enter username (Github username) and password (Token generated first step)
11. Click Finish to start the cloning process
13. Right click project (SLA_report2) once cloned, click Build Path -> Configure Build Path -> Libraries -> Add Jars, then add all jar external files saved in file named "Jar Files" in this repo then click Apply and Close
14. Download NEW-IT Contractor-VG-Vendor Req Report and Resource Roster to local machine (I recommend saving both to Desktop for easy path configuration)
15. Click run (circular green play button on top left of eclipse once project, SLA_report2, has been selected)
16. Enter the inputs instructed and file should be created based on the path you provide
17. To configure a path of both files: right click on the file, click properties then combine "Location", "name", and ".xlsx" together (ie. Location: C:\User\name\Desktop and Resource Roster would be C:\User\name\Desktop\Resource Roster.xlsx)
18. Repeat for where you want to write your file to (ie. if you want to save to desktop, input path C:\User\name\Desktop\Entech IT Staff MON XQ20XX.xlsx)
19. Once program completes, file will be saved to path you specified 
