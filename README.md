# Test-Case-Migrator-Plus
Test Case Migrator Plus tool allows test related artifacts, present in Excel and MHT/Word formats, to be imported into Team Foundation Server. It works with both <b> Visual studio 2017, Visual studio 2019; and TFS 2017, TFS 2018 and Azure DevOps Server 2019.1 &amp; 2019.1+ </b>

<b> No changes to the license, Microsoft public license.</b>

<a href="https://github.com/premboyapati/Test-Case-Migrator-Plus/tree/master/Binaries"> Downalod Binaries Here </a>

QA team in my company always uses Test case migrator tool to import test cases from Excel. There was no issue of using TCM import tool when my company using TFS 2010/2012.

TCM import tool suddenly "stopped" when my company recently moved to TFS 2017 and then moved Azure Devops. QA team was trying to resolve the issue themselves, talking to people in differen Microsoft/other forums. <b> They spent months but no luck. And also they were waiting for Microosft release new version which works with TFS 2017+ </b>

As a last hope they came to me asking help to upgrade this TCM import tool to make it work with Azure DEVOps. 

I wanted to help them by sharing updated binaries by downloading patches and sourcode from codeplex "https://www.sharepointpals.com/post/how-to-migrate-the-test-cases-from-an-excel-to-tfs/"

But no luck with the above Url, it always redirects archives, that is when I decided to upgrade the source code to work with .Net framweork 4.6.1.

After upgrading TCM import source code with .Net4.6.1, build is not compiled and got 456 compilation errors. To fix all these issues, I modified code and took almost half a day to make it work.

Now, my team and other QA teams in my company are pretty happy with this updated binaries. <b>And that's when my team asked me to upload this binaries and source code to GitHub to share with the QA world to use it in your day to day job.</b>

Please downaload the <a href="https://github.com/premboyapati/Test-Case-Migrator-Plus/tree/master/Binaries"> binaries </a> and use the tool. 

<i> I or my team or GitHub or Microsoft is not repsonsible for any issues if you face while using this product.You are sole responsible in using this product. My role ends in getting the code from codeplex and converting it to make it work with TFS 2017+ </i>


