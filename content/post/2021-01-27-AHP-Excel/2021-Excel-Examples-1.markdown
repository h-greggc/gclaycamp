---
title: "Spreadsheets for Quality Risk Management?"
author: "Gregg Claycamp"
date: '2021-01-13'
output:
  html_document:
    toc: true
    df_print: paged
  pdf_document: default
  word_document: default
subtitle: ''
table_of_contents: true
slug: Excel-examples
summary: Microsoft Excel<sup>TM</sup> is perhaps the most common spreadsheet software on business desktops. Often it is the main tool for small businesses and a prototyping
  and team- or division-based tool for larger business. Spreadsheet software offers
  an easy way to control and automate repetitive processes from data collection to
  knowledge management. Finally, Excel macros are a convenient framework for teaching
  basic methods in decision analyis. Given its pros and cons, a spreadsheet can be
  a useful tools for quality risk management operations and decisions.
lastmod: '2021-01-13'
featured: yes
image:
  caption: 'Image credit: [**John Moeses Bauan**](https://unsplash.com/photos/OGZtQF8iC0g)'
  placement: 3
header:
  image: null
categories:
- QRM
- MCDA
- Decision Analysis
- Excel
draft: yes
---

<script src="/rmarkdown-libs/kePrint/kePrint.js"></script>
<link href="/rmarkdown-libs/lightable/lightable.css" rel="stylesheet" />
<script src="/rmarkdown-libs/htmlwidgets/htmlwidgets.js"></script>
<script src="/rmarkdown-libs/jquery/jquery.min.js"></script>
<link href="/rmarkdown-libs/dygraphs/dygraph.css" rel="stylesheet" />
<script src="/rmarkdown-libs/dygraphs/dygraph-combined.js"></script>
<script src="/rmarkdown-libs/dygraphs/shapes.js"></script>
<script src="/rmarkdown-libs/moment/moment.js"></script>
<script src="/rmarkdown-libs/moment-timezone/moment-timezone-with-data.js"></script>
<script src="/rmarkdown-libs/moment-fquarter/moment-fquarter.min.js"></script>
<script src="/rmarkdown-libs/dygraphs-binding/dygraphs.js"></script>

## Still Popular, after all these years…

Although data scientists and statisticians might groan at data automation and analysis using spreadsheet platforms, such as Excel<sup>TM</sup>, Google Sheets<sup>TM</sup> or others, the long-standing popularity of spreadsheet tools remains indisputable. Beginning with simple monochrome tables using Visicalc in 1979, the spreadsheet has evolved to a user-friendly interface for accounting tables, diverse analyses and dashboards for summaries of multiple lines sources of data. This fact and widespread availability in several formats keep the software popular for personal, business team and even enterprise IT solutions. Some reasons for the popularity of spreadsheets include their

-   accessibility to non-data scientists and non-programmers;

-   availability on most platforms including Windows, iOS, Android and cloud servers;

-   portability–the same spreadsheet can be run independently on a laptop or launched on the internet;

-   cell formulas, many of which are intuitive for basic math needs; and

-   support from a large online community of spreadsheet users, consultants and software vendors.

<figure>
<img src="/img/2021-01-27-AHP-Excel/Visicalc.png" data-position="center" width="600" alt="Fig. 1. Screenshot of Visicalc in 1979. About 20 rows in 4 columns were all that could be viewed in the typical, small monochrome monitor!" /><figcaption aria-hidden="true"><strong>Fig. 1. Screenshot of Visicalc in 1979. About 20 rows in 4 columns were all that could be viewed in the typical, small monochrome monitor!</strong></figcaption>
</figure>

A full-featured application, such as Excel, can also bridge many organizational needs from simple maintenance of a contacts list to an informative dashboard backed by automated data analysis (e.g., [TheSmallman.com](https://www.thesmallman.com/)). It is apparent that some industries see a positive trend in using Excel dashboards to monitor key process indicators (KPIs) for the organization even if the dashboard is only a GUI for data from more sophisticated databases and statistical software.

<figure>
<img src="/img/2021-01-27-AHP-Excel/HR-Dashboard.jpg" data-position="center" width="600" alt="Fig.2. A Random Example of a Dashboard" /><figcaption aria-hidden="true"><strong>Fig.2. A Random Example of a Dashboard</strong></figcaption>
</figure>

Given the wide utility of (programmable) spreadsheets, motivated spreadsheet users can expect to find both room for growth and ample support developing deeper expertise using such tools. The remainder of this post will some limitations and uses of Excel for quality risk management (QRM) and QRM decision-making.

> Excel is not going anywhere, and businesses will continue to use Excel as a primary tool for diverse functions and applications ranging from IT projects to company picnics. [Investopedia June 25, 2019](https://www.investopedia.com/articles/personal-finance/032415/importance-excel-business.asp)

### Limitations to spreadsheet software

Given all the positives, why are data scientists, engineers and other “quants” so negative about spreadsheet tools? Many of the blog posts compare Excel or other spreadsheets with statistical software such as [R](https://www.r-project.org/about.html), [SAS](https://www.sas.com "Statistical software - SAS"), [Tableau](https://www.tableau.com/) or [MiniTab](www.minitab.com). But these comparisons are misleading because they are “apples-to-oranges” comparisons that do not often reflect the different types and needs of users. Although spreadsheets have impressive lists of built-in and add-in statistical functions, the spreadsheet platforms are far from being pre-programmed for descriptive, inferential and other statistical modeling that is found in advanced statistical packages. Second, rigorous quantitative solutions for business analysis and forecasting can be programmed in [Visual Basic for Applications](https://docs.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office), C\# or C<sup>++</sup>, [Javascript](javascript.com), or [Typescript](https://www.typescriptlang.org/), all launched “behind” Excel. But compilation and deployment of programs in these languages tools can be daunting for the typical spreadsheet user and macro programmer. Third, program modules are difficult to manage in an enterprise IT organization and, (albeit in a shrinking number of cases), unsafe to share across the internet. Finally, spreadsheet programs typically run in a size-limited container making them unwieldy or even useless for a “big data” project. For example, the worksheet limit is 1,048,576 rows, although there are work-arounds using Excel’s data model tools (e.g, see: [How can You Analyze 1 mn Rows of Data - Excel Interview Question - 02](https://chandoo.org/wp/more-than-million-rows-in-excel/) and below).

<iframe width="549" height="396" frameborder="0" scrolling="no" src="https://onedrive.live.com/embed?resid=B16B51D1F70656AA%2169565&amp;authkey=%21AIVoxt7YVkR39Gk&amp;em=2&amp;wdAllowInteractivity=False&amp;Item=Chart%203&amp;wdDownloadButton=True&amp;wdInConfigurator=True">
</iframe>

## Spreadsheets in Quality Risk Management

Although QRM is a relatively new concept and practice, already there are numerous Excel and other spreadsheet apps, templates and add-ins that are applicable to QRM functions. For example, a Google search for “quality risk management” AND “Excel template” yields about 19,400 results![^1] Given that QRM is a multidisciplinary blend from quality assurance, risk management and decision analysis, it is likely that most QRM spreadsheet apps come from modifications of the spreadsheets apps and templates in those disciplines. In essence, QRM spreadsheet apps and templates are a small subset of a vast set for more general quality management purposes.[^2]

### Comparing Spreadsheets to R and Python

Perhaps the two most popular tools for data science are the open-source languages [R](https://www.r-project.org/about.html) and [Python](https://www.python.org/). R is a statistical language program that is both free and easy to learn. Python is a free programming language that has a significant set of packages for data analysis and scientific computing. Both programs can “run anywhere” and they can be used with user-friendly integrated development environments (IDEs) such as [RStudio](rstudio.com) for R and [Idle](https://docs.python.org/3/library/idle.html) for python. For example, this blog and website are developed in RStudio using the R package, [blogdown](https://bookdown.org/yihui/blogdown/).

In spite of apparent unpopularity for statistical programming and data chores, desktop spreadsheet software remains popular due to its broad accessibility to users. Not only are spreadsheet programs available on almost every desktop, laptop and tablet produced, but the documentation and training opportunities for beginning to advance users are ubiquitous. re

Others’ opinion: <https://martinctc.github.io/blog/my-favourite-alternative-to-excel-dashboards/#fn1>

vs Python [Excel vs python](https://www.nobledesktop.com/learn/python/python-vs-excel)

Use python not excel :<https://www.dataquest.io/blog/excel-vs-python/>

[R versus Excel: What’s the Difference](https://www.northeastern.edu/graduate/blog/r-vs-excel/)

[Understanding R programming over Excel for Data Analysis](https://www.gapintelligence.com/blog/understanding-r-programming-over-excel-for-data-analysis/)

Alternatives: DT [DataTables in the DT package](https://rstudio.github.io/DT/)

flexDashboard R package [flexDashboard](https://rmarkdown.rstudio.com/flexdashboard/)

excel not even mentioned in survey of BI tools: <https://bi-survey.com/business-intelligence-software-comparison>

``` r
library(dygraphs)
dygraph(nhtemp, main = "New Haven Temperatures") %>% dyRangeSelector(dateWindow = c("1920-01-01", 
    "1960-01-01"))
```

<div id="htmlwidget-1" style="width:672px;height:480px;" class="dygraphs html-widget"></div>
<script type="application/json" data-for="htmlwidget-1">{"x":{"attrs":{"title":"New Haven Temperatures","labels":["year","V1"],"legend":"auto","retainDateWindow":false,"axes":{"x":{"pixelsPerLabel":60}},"showRangeSelector":true,"dateWindow":["1920-01-01T00:00:00.000Z","1960-01-01T00:00:00.000Z"],"rangeSelectorHeight":40,"rangeSelectorPlotFillColor":" #A7B1C4","rangeSelectorPlotStrokeColor":"#808FAB","interactionModel":"Dygraph.Interaction.defaultModel"},"scale":"yearly","annotations":[],"shadings":[],"events":[],"format":"date","data":[["1912-01-01T00:00:00.000Z","1913-01-01T00:00:00.000Z","1914-01-01T00:00:00.000Z","1915-01-01T00:00:00.000Z","1916-01-01T00:00:00.000Z","1917-01-01T00:00:00.000Z","1918-01-01T00:00:00.000Z","1919-01-01T00:00:00.000Z","1920-01-01T00:00:00.000Z","1921-01-01T00:00:00.000Z","1922-01-01T00:00:00.000Z","1923-01-01T00:00:00.000Z","1924-01-01T00:00:00.000Z","1925-01-01T00:00:00.000Z","1926-01-01T00:00:00.000Z","1927-01-01T00:00:00.000Z","1928-01-01T00:00:00.000Z","1929-01-01T00:00:00.000Z","1930-01-01T00:00:00.000Z","1931-01-01T00:00:00.000Z","1932-01-01T00:00:00.000Z","1933-01-01T00:00:00.000Z","1934-01-01T00:00:00.000Z","1935-01-01T00:00:00.000Z","1936-01-01T00:00:00.000Z","1937-01-01T00:00:00.000Z","1938-01-01T00:00:00.000Z","1939-01-01T00:00:00.000Z","1940-01-01T00:00:00.000Z","1941-01-01T00:00:00.000Z","1942-01-01T00:00:00.000Z","1943-01-01T00:00:00.000Z","1944-01-01T00:00:00.000Z","1945-01-01T00:00:00.000Z","1946-01-01T00:00:00.000Z","1947-01-01T00:00:00.000Z","1948-01-01T00:00:00.000Z","1949-01-01T00:00:00.000Z","1950-01-01T00:00:00.000Z","1951-01-01T00:00:00.000Z","1952-01-01T00:00:00.000Z","1953-01-01T00:00:00.000Z","1954-01-01T00:00:00.000Z","1955-01-01T00:00:00.000Z","1956-01-01T00:00:00.000Z","1957-01-01T00:00:00.000Z","1958-01-01T00:00:00.000Z","1959-01-01T00:00:00.000Z","1960-01-01T00:00:00.000Z","1961-01-01T00:00:00.000Z","1962-01-01T00:00:00.000Z","1963-01-01T00:00:00.000Z","1964-01-01T00:00:00.000Z","1965-01-01T00:00:00.000Z","1966-01-01T00:00:00.000Z","1967-01-01T00:00:00.000Z","1968-01-01T00:00:00.000Z","1969-01-01T00:00:00.000Z","1970-01-01T00:00:00.000Z","1971-01-01T00:00:00.000Z"],[49.9,52.3,49.4,51.1,49.4,47.9,49.8,50.9,49.3,51.9,50.8,49.6,49.3,50.6,48.4,50.7,50.9,50.6,51.5,52.8,51.8,51.1,49.8,50.2,50.4,51.6,51.8,50.9,48.8,51.7,51,50.6,51.7,51.5,52.1,51.3,51,54,51.4,52.7,53.1,54.6,52,52,50.9,52.6,50.2,52.6,51.6,51.9,50.5,50.9,51.7,51.4,51.7,50.8,51.9,51.8,51.9,53]]},"evals":["attrs.interactionModel"],"jsHooks":[]}</script>

[^1]: Retrieved 3/16/2021. This search does not filter for multiple hits on a single website.

[^2]: For example, a search of “spreadsheets for quality assurance” receives 2,790,000 hits! (Retrieved March 15, 2021).
