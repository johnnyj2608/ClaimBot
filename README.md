## Project Name & Pitch

Claimbot

A Python script with a GUI to automatically submit health insurance claims on Office Ally

3 main components: Selenium, CustomTKinter, and XlWings

1. Selenium: Automates the login process and auto-fills health insurance claims. Utilized explicit waiting to ensure elements were present before sending keys.

2. CustomTkinter: User-friendly GUI to select options for automating claims. Utilized threading to allow concurrent use of the GUI to "stop automation."
  
3. XlWings: Stores data and allows easy updates to member or default values of claims. Uses Pandas to extract the information and convert the information to dictionaries. After each claim is submitted, the claim date, member name, range of dates, and amount will be recorded on the Claims tab.

### Demo Video
[![Watch the demo](https://img.youtube.com/vi/kw0KtUPqezA/0.jpg)](https://www.youtube.com/watch?v=kw0KtUPqezA)

## How To Use

### Excel values must be identical to stored info (capitalization, [123], etc.)

0. Acquire a copy of the Claimbot Template. Select the relevant claim form, provide values to the fields, and write your member's information
![1](https://github.com/johnnyj2608/ClaimBot/assets/54607786/7660badc-427c-4d1f-a47d-e8996c7d64e4)
![2](https://github.com/johnnyj2608/ClaimBot/assets/54607786/7a028d5b-1ee5-4ddd-bba4-cb5a8bd9d96d)

1. Select the Excel template to retrieve information from
2. Select a range of members and dates to write claims for
3. Optional checkbox for automatically submitting and downloading PDFs
4. Let the script go to work
5. All submitted claim information will be written to the Excel file

![3](https://github.com/johnnyj2608/ClaimBot/assets/54607786/e42c58be-4ed3-49dc-9008-a264edb043f3)

**Vacation** allows you to remove a range of dates from the submission process.  
**Excluded** is used if you wish to remove the member from the automation process without deleting it from Excel.

Pyinstaller: pyinstaller main.spec
Executable will be located in the dist directory

## Reflection

The project came to me when I watched my coworker tediously point and click at health insurance forms, taking about 5 minutes per claim and another 5 to rest between each. Over 400 claims would need to be submitted biweekly. With my knowledge of programming, I told my coworker "Let me see what I can do."

Since the core of this project was auto-fill, I was debating between a Javascript (Google Chrome Extension) or Python (Selenium). Eventually, I settled for Python because using an Excel file for data storage was a necessity due to coworker constraints. I previously had experience with the Python library XlWings (Excel) and CustomTKinter (GUI). In researching the best automation tools, Selenium seemed to be the best fit for the scope of the project. I watched 2 Selenium YouTube tutorials to understand how to use it and then started working on the project.

There were complications along the way. The biggest headscratcher I had was "iframes". Sometimes elements are embedded within an "iframe", which makes buttons, text fields, etc. non-interactable. Eventually, I learned that the driver must switch to the new iframe before it works. 

Another difficulty I had was getting the "automate download" feature to work. OfficeAlly "print" button does not start the download of the PDF but rather displays the Google Chrome PDF viewer, which Selenium can not interact with. Luckily I noticed that the PDF link is associated with the claim ID, which I was able to retrieve. I then would use the driver to get the PDF link and download it that way.

Finally, integrating a list box with a search entry seemed straightforward initially, but maintaining multiple selection highlights between search filters proved challenging. Without a built-in method to get line numbers for selection/deselection, I first tried saving all the line numbers, only to find indices changed with each search filter. Switching to saving actual values in sets resolved that problem but then I learned set intersections didn't account for filtered search results. After much frustration and giving up twice, I finally realized I could use the list of members to remove filtered results, allowing set intersections to work correctly. Despite the hurdles, spending time off the computer and sleeping on the idea eventually led to a successful implementation, proving the effort worthwhile.

After testing and receiving feedback from my coworker, the program is ready for deployment. The speed at which health insurance claims are being submitted is 80% faster. Currently, this project is only usable for OfficeAlly.com (Classic). Still, it can easily scale to other health insurance websites like Availity by including another field that selects the relevant platform for the insurance so it would run the correct Python script.

## Project Screen Shots

![1](https://github.com/johnnyj2608/ClaimBot/assets/54607786/e46ad759-cc4b-47f7-8f35-c3c25c55f9a6)
![2](https://github.com/johnnyj2608/ClaimBot/assets/54607786/4e2254fd-07cf-4016-a851-327803b3bc6c)

