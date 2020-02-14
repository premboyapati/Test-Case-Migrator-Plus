_No changes to the license, Microsoft public license._

Test Case Migrator Plus tool allows test related artifacts, present in Excel and MHT/Word formats, to be imported into Team Foundation Server. It works with both _Visual studio 2017, Visual studio 2019; and TFS 2017, TFS 2018 and Azure DevOps Server 2019.1 & 2019.1_

**For more details: https://archive.codeplex.com/?p=tcmimport**

1. Can not run existing TCM import EXE
2. No use of this patch: https://tcmimport.codeplex.com/SourceControl/list/patches. 
3. Are you looking for a tool to upload all the test cases to TFS 2017, 2018, 2019 and Azure DEVOps? 
4. Is Test case Migration tool downloaded from Codeplex (or codeplex archives) is not working?

**_Do not worry, here you go latest, tested and working binaries are here [TCM Plus](https://github.com/premboyapati/Test-Case-Migrator-Plus/tree/master/Binaries) to rescue you._**
***

### Prerequisites:
 * .Net framework 4.6.1 or 
 * VS 2017 or 
 * VS 2019.
***

# Story behind bringing this code from Codeplex to GitHub
QA team in my company always uses Test case migrator tool to import test cases from Excel. There was no issue of using TCM import tool when my company using TFS 2010/2012.

TCM import tool suddenly **stopped** running when my company recently moved to TFS 2017 and then moved Azure Devops. QA team was trying to resolve the issue themselves, talking to people in differen Microsoft/other forums. They spent months but no luck. And also they were waiting for Microosft release new version which works with TFS 2017+

As a last hope they came to me asking help to upgrade this TCM import tool to make it work with Azure DEVOps.

I wanted to help them by sharing updated binaries by downloading patches and sourcode from codeplex "https://www.sharepointpals.com/post/how-to-migrate-the-test-cases-from-an-excel-to-tfs/"

But no luck with the above Url, it always redirects archives, that is when I decided to upgrade the source code to work with .Net framework 4.6.1.

After upgrading TCM import source code with .Net4.6.1, build is not compiled and got 456 compilation errors. To fix all these issues, I modified code and took almost half a day to make it work.

Now, my team and other QA teams in my company are pretty happy with this updated binaries. And that's when my team asked me to upload this binaries and source code to GitHub to share with the QA world to use it in your day to day job.

Please [downlaod](https://github.com/premboyapati/Test-Case-Migrator-Plus/tree/master/Binaries) the binaries and use the tool.

_I or my team or GitHub or Microsoft is not responsible for any issues if you face while using this product.You are sole responsible in using this product. My role ends in getting the code from Codeplex and converting it to make it work with TFS 2017+_
