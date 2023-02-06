<h1 align="center">Plano Basico's Webscraping Automation 🤖</h1>

<p align="center">
 <a href="#description">Description</a> •
 <a href="#features">Features</a> • 
 <a href="#code_steps">Code Steps</a> •
 <a href="#requirements">Requirements</a> • 
 <a href="#author">Author</a>
</p>

<h4 align="center"> 
	✅  Python WebScraping Automation Project 🚀 Finished...  ✅
</h4>

<p></p>

Description
=======================
<p align="center">A webscraping automation thats go to the desired page, get all filtered data, treat, then saves on an excel's file.</p>

<a href="https://www.python.org/">
    <img src="https://img.shields.io/static/v1?label=Language&message=Python&color=3776AB&style=for-the-badge&logo=ghost" /> . 
</a>
<a href="https://pandas.pydata.org/">
    <img src="https://img.shields.io/static/v1?label=Library&message=Pandas&color=150458&style=for-the-badge&logo=ghost" /> .
</a>
<a href="https://www.selenium.dev/">
    <img src="https://img.shields.io/static/v1?label=Library&message=Selenium&color=43B02A&style=for-the-badge&logo=ghost" />
</a>
<p></p>

Features
=======================
- [x] Go to desired page
- [x] Get all data from it
- [x] Saves as [pandas](https://pandas.pydata.org/) dataframe
- [x] Manipulate and process data
- [x] Saves as an excel file

Code_Steps
=======================
1. The code go to desired page (url)
2. It show just 250 itens per page, so is needed to set this on the table's page view settings
3. Filter the 'Serviço' field with 'TV' name
4. As there are 250 itens per page, one needs to know how many pages will need to be viewed to get all the data
5. Every page is a new dataframe, so we concat the old dataframe with the new until we see all the pages
6. After getting all data, we manipulate and process data with some custom rules
7. By the end, we save all as an excel file

## Preview
<p align="center">
  <img alt="Preview Python WebScraping Automation Project" src="https://github.com/J-Neto/plano-basico-webscraping-automation/blob/master/preview-excel-sheet.png">
</p>

Requirements
=======================
Before starting, you must have installed the following tools on your machine: [GIT](https://git-scm.com/downloads), [PYTHON](https://www.python.org/downloads/), [EXCEL](https://www.microsoft.com/pt-br/microsoft-365/excel). Also it's nice to have an editor to work with the code like [VSCode](https://code.visualstudio.com/).

Also, you must have the following libraries:
```bash
* selenium
* webdriver_manager
* pandas
* lxml

# Install them with this command
$ pip install <library>
or
$ pip3 install <library>
```

```bash
# Clone this repository
$ git clone https://github.com/J-Neto/plano-basico-webscraping-automation.git

# Access project's folder in terminal/cmd
$ cd plano-basico-webscraping-automation

# Open the code in the editor
$ code .
or manually open via file explorer by mouse right clicking or by VSCode on the path "File > Open Folder > [Find the folder "plano-basico-webscraping-automation"] > Open"

# Run python's file
$ python plano-basico-webscraping-automation.py
```

Author
=======================
<a href="https://https://github.com/J-Neto"><img src="https://avatars.githubusercontent.com/u/49914443?v=4" width="100px;" alt=""/><br><p><b>José Neto</b> 🎧</p></a><br>

Made with ❤️ by José Neto 👋🏽 Get in touch! </br>
Feel free to ask me anything, I'll be glad to help 😉

<a href="https://www.linkedin.com/in/jos%C3%A9-neto-299920152/"> <img src="https://img.shields.io/badge/LinkedIn-%230A66C2?style=for-the-badge&logo=linkedin&logoColor=white" alt="LinkedIn"></a> 
<a href="mailto:ribeirojoseph44@gmail.com"> <img src="https://img.shields.io/badge/Gmail-%23C5221E?style=for-the-badge&logo=gmail&logoColor=white" alt="Gmail"></a> 
<a href="https://www.instagram.com/neto._ribeiro/"> <img src="https://img.shields.io/badge/instagram-%23FE2973.svg?&style=for-the-badge&logo=instagram&logoColor=white" alt="Intagram"></a>