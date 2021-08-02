# MKPIS

**MADE MY (ZackAllen1) OWN MODIFICATIONS FOR USE WITH GSM**

**MODIFICATIONS INCLUDE**
- Converting duration format
- Added Export to CSV using Excelize Package
- Routed CSV To Power BI Report for consumption
- Added YAML build to test Github Token Configuration

**HOW TO RUN**
- Go to [https://github.com/ZackAllen1/mkpis](https://github.com/ZackAllen1/mkpis)
- Click on the Green Code button and click Download ZIP
- Unzip the file and open the folder in an IDE that can run Go code (Ex: VSCode)
  * If using VSCode, upon opening the file you can install the Go Extension 
- Change the directory to the following using the Terminal
  ```
  \mkpis\cmd\cli
  ```
- Create a file called .env in this directory and add the following code
  ```
  GITHUB_TOKEN= <Your Token Here>
  DEVELOP_BRANCH_NAME= <Feature Branch Name> 
  MASTER_BRANCH_NAME= <Master Branch Name>
  ```
- To get the GITHUB_TOKEN, do the following steps:
  * Go to [https://github.com/settings/tokens](https://github.com/settings/tokens)
  * Click Generate New Token
  * Set an Expiration Date (Note: you will have to repeat this process to regenerate the token after the expiration date)
  * Check all of the checkboxes and click the Green Generate Token Button
  * Copy the token that was created and replace ```<Your Token Here>``` with the token you just created

- ```<Feature Branch Name>``` generates the following data about the branch
  ![feature branch diagram](featurebranch.png)
- Wheras ```<Master Branch Name>``` generates the following data about the branch
  ![master branch diagram](masterbranch.png)


- Save the .env file
- Open the terminal again (making sure the end of the path is ```mkpis\cmd\cli```) and use the following command
  ```
  go run main.go -owner GlobalSafetyManagement -repo <RepoName> -from "YYYY-MM-DD" -to "YYYY-MM-DD"
  ```
  * ```<RepoName>``` can be any repository within GlobalSafetyManagement (Ex: SDS-Application, AutoSDS, etc.)
  * ```-from``` is the start date of the report and ```-to``` is the end date of the report
- Example
  ```
  go run main.go -owner GlobalSafetyManagement -repo SDS-Application -from "2021-05-01" -to "2021-07-30"
  ```
  * Will run the report on SDS-Application from May 1st 2021 to July 30th 2021
- Once you run the command a file called ```GithubBranchReport.xlsx``` will be created
  * This file is what will be used in the Power BI Report



**Importing to Power BI**
- Open the GSM Workspace on the Power BI Website
- Click on the GithubPRTracker Report
- Click File > Download the .pbix File
- Open Power BI Desktop and load the .pbix File you just downloaded
- On the Home bar click the drop down arrow on "Transform Data", then click "Data Source Settings"
- Make sure "Data sources in current file" is selected, then "Change Source" to the path of ```GithubBranchReport.xlsx``` that was created on your desktop
- Close then Refresh, and the data should update
- Now anytime you run the ```go run main.go``` command from above, all you have to do is refresh the Power BI data and "Publish" your data for everyone to see


## END OF NOTES



Measuring the development process for Gitflow managed projects.


## Visibility over the process 

Part of a healthy engineering culture is to foster team and developers to continuously improve their work processes.

In order to improve the Engineering Managers/Software Managers/TechLeads need visibility over the development process.

Of course, there are tons of metrics you can extract from the whole development process, since the inception until the code reaches production, many companies rely on the agile definition of Team velocity or on some metrics you can easily extract from Jira.

But with those metrics still have some blind spots when we are talking about the development process.

This project put specific focus on the feature pull request and the release pull request and from those extract as much information it can to help to measure, understand and improve the collaboration and engagement in your team and culture.

Howbeit, you shouldn't use these metrics to compare teams or individuals rather than debug and improve the process.

The numbers can change in the order of magnitude depending on the kind of project or the kind of task the developer is working on.

## Software Engineering Metrics 

Team/developer effort can’t be measured just by how many pull requests are merged or how big the PR is.

This application will provide a bunch of metrics but you should give them meaning align with your organization engineering culture.

For example, fast code reviews are a good sign, but if see them on big tasks and with quite a long time until the revision starts its a clear sign that the developers are afraid or they are lazy to review 500 lines and they just skipped. 

Another example, a long time before and after the CR it can be a signal of a lot of issues that you should investigate for instance it can be a signal of constant merging conflicts if it goes together long times between the first commit until the merge. 

Your team is complex and the secrets to debug and improve it are in the details rather than in just one metric.

You can also use these metrics to test your theories and prove the changes you applied are the correct ones.

How can I decrease more the QA time, hiring another QA or forcing the developers to test at least the happy path?

Is it true that doing the task in peer programming will reduce the QA time and remove the Code review time? is it worth it?


## How does it work?


This application heavily realy on Github API (more providers in the roadmap) to extract information about the merged pull request given a window of time.

Once it has all the necessary data it builds a report in your terminal, with all feature request and releases metrics.

**Command:**

`mkpis -owner RepoOwner -repo RepoName`

**Usage:**

<pre>
  -from string
        When the extraction starts (default "2020-08-15")
  -owner string
        Owner of the repository
  -repo string
        Repository name
  -to string
        When the extraction ends (default "2020-08-25")
</pre>

**Example**

![Example screencast](docs/mkpis.gif)

**Enviroment variables**

To run this application is mandatory to have a GitHub token with the right permission in your environment.

`GITHUB_TOKEN=XXXXXXXXXXXXXXXXXXXXXXXXXXXX`

You can customize the master and development branch names via enviroment variables:

*.env*
<pre>
DEVELOP_BRANCH_NAME=devel
MASTER_BRANCH_NAME=master
</pre>

**Note:** The application automatically reads *.env* files in the execution path.
 


## Pull request KPI



### Feature Pull Request

<pre>
+----------+              +------+               +--------------+                 +-------------+               +-------+
|          |    Opens     |      |  Waits for    |              |  Discuss until  |             |   Waits for   |       |
|  Commit  +-------------->  PR  +---------------> First Review +-----------------> Last Review +---------------> Merge |
|          |              |      |               |              |                 |             |               |       |
+----------+              +------+               +--------------+                 +-------------+               +-------+

+                                              Feature Lead Time                                                        + 
+-----------------------------------------------------------------------------------------------------------------------+
+                                                                                                                       +
                                                               Pull Request Lead time
                          +                                                                                             +
                          +---------------------------------------------------------------------------------------------+
                          +                                                                                             +
                                                                 Review time
                                                 +                                              +
                                                 +----------------------------------------------+
                                                 +                                              +
                            Time to first review                                           Last review to merge time
                          +                      +                                  +                                   +
                          +----------------------+                                  +-----------------------------------+
                          +                      +                                  +                                   +

</pre>

**Feature Lead Time:** it measures how much time the first commit in the pull request takes to reach the devel branch.

`Formula: (merged_at  - first_commit_created_at)`

**Pull Request Lead time:** it measures the time the code review process plus the corrections needed to move merge the code. It considers all the steps of code review and the wating times.

`Formula: (merged_at – opened_at)`

**Review time:** it measures the time from the first to the last review in the pull request.
`Formula: (last_review_created_at  – first_review_created_at )`

**Time to First Review:** it measures how much time the team takes to make the first code review. 

`Formula: (first_review_created_at - last_commit_created_at)`

**Last Review to Merge Time:** it measures hoy much time the feature remain unmerged after the last review.

`Formula: (merged_at - last_commit_created_at)`

**Pull request Size:** it measures the pull request size in terms of changes it contains.


`Formula: (pull_request_additions + pull_request_deletions)`


### Release Pull Request

<pre>
+----------+              +------+                  +-------+
|          |    Opens     |      |  Waits for QA    |       |
|  Commit  +-------------->  PR  +------------------+ Merge |
|          |              |      |                  |       |
+----------+              +------+                  +-------+

+                     Release Lead Time                     +
+-----------------------------------------------------------+
+                                                           +
                                Release Review Lead time
                          +                                 +
                          +---------------------------------+
                          +                                 +
</pre>

**Release Lead Time:** it measures how much time the first commit in the release request takes to reach the master branch. It's quite common tu assume this metric as *from code to deploy time*.

`Formula: (merged_at  - first_commit_created_at)`

**Release Review Lead Time:** it measures how much time the release review takes, in a lot of organizations it means QA time.

`Formula: (merged_at – opened_at)`

**Relase Size:** it measures the release size in terms of changes it contains.


`Formula: (pull_request_additions + pull_request_deletions)`


## Limitations

* Currently this application only work in github repos.
* This project is useless if your team is working in a Trunk Base development.



## Licence

Copyright © 2020, Jordi Martín (http://jordi.io)

Released under MIT license, see LICENSE for details.

