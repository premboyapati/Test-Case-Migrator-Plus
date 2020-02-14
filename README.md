
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
<p>QA team in my company always uses **Test case migrator tool** to import test cases from _Excel to TFS_. It was a _perfect tool_ for our QA team to load test case to TFS when my company using TFS 2010/2012.

TCM import tool suddenly **stopped** working when my company upgraded to Azure Devops. QA team was trying to resolve the issue themselves, talking to people in Microsoft/other forums. 
They **spent** months but **no luck**. And also they were waiting for **Microsoft** (_My team told me Microsoft is still working on upgrading this tool_) to release new version which works with TFS 2017+.

As a last hope they came to me asking help to upgrade this TCM import tool to make it work with Azure DEVOps.
</p>
<br /><br /><br />
I wanted to help them by **sharing ** updated binaries by downloading patches and source code from Codeplex "https://www.sharepointpals.com/post/how-to-migrate-the-test-cases-from-an-excel-to-tfs/".

But **no luck** for me with the above Url, it always **redirects **to archives folder, that is when I decided to **upgrade **the source code using **.Net framework 4.6.1** to make it to work with **TFS 2017+.**

After upgrading TCM import source code with .Net4.6.1, build is **not **compiled and got **400+ **compilation errors. To fix all these issues, I **modified **code, code changes took almost _half a day_ to make it work.

<br /><br /><br />
Now, my team and other QA teams in my company are pretty happy (_I would say there are on cloud nine :)_ ) with this updated binaries. And that's when my team asked me to u_**pload this binaries and source code**_ to GitHub to share with the **QA world** to use it in your day to day job.

Please [downlaod](https://github.com/premboyapati/Test-Case-Migrator-Plus/tree/master/Binaries) the binaries and use the tool.

_I or my team or GitHub or Microsoft is not responsible for any issues if you face while using this product.You are sole responsible in using this product. My role ends in getting the code from Codeplex and converting it to make it work with TFS 2017+_
