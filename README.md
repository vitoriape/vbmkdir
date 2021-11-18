<h3 align="center"> 
<img alt="vmkdir banner" src="./images/vbmkdir.banner.png" width="1000" height="200">
</h3>

<p align="center">
  <img alt="GitHub language count" src="https://img.shields.io/github/languages/count/vitoriape/vbmkdir">
  
  <img alt="GitHub top language" src="https://img.shields.io/github/languages/top/vitoriape/vbmkdir">
  
  <a href="https://github.com/vitoriape/vbmkdir/blob/mkdir.vb-vpa/LICENSE">
    <img alt="License: MIT" src="https://img.shields.io/badge/License-MIT-green.svg">
  </a>
  
  <a href="https://github.com/vitoriape/vbmkdir/commits/master">
    <img alt="GitHub last commit" src="https://img.shields.io/github/last-commit/vitoriape/vbmkdir">
  </a>
</p>

---

Index
=================
<!--ts-->
   * [About](#about)
   * [Tools](#tools)
   * [References](#references)
   * [Setup](#setup)
   * [Author](#author)
 

### About

> [Versão em Português (pt-br)](https://github.com/vitoriape/vbmkdir/blob/master/LEIAME.md). <img src="https://camo.githubusercontent.com/dcc375ada213d3ac04a9781518098cd4d071601bc2ccfc120025cc32b6d38fab/68747470733a2f2f63646e2e737461746963616c792e636f6d2f67682f686a6e696c73736f6e2f636f756e7472792d666c6167732f6d61737465722f7376672f62722e737667" alt="brazil flag" width="20" height="20">

This project is an [VBA](https://docs.microsoft.com/pt-br/office/vba/library-reference/concepts/getting-started-with-vba-in-office) script for create folders automatically from sellected cells on Excel.


### Tools

Development of this script utilizes the tools listed below:

- [Git](https://git-scm.com/)
- [Excel](https://support.microsoft.com/en-us/excel)
- [Visual Basic for Applications](https://docs.microsoft.com/pt-br/office/vba/api/overview/excel)


### References

For more information about using the statement `Do (...) Loop` and statement `MkDir`, besides the function `Dir` on Visual Basic for Applications, read the Microsoft documentation:

- [Do...Loop Statement](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/doloop-statement)
- [MkDir Statement](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/mkdir-statement)
- [Dir Function](https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dir-function)


### Setup

```cmd
# Clone this repository
$ git clone <https://github.com/vitoriape/vbmkdir>
```

* <b>1. Make sure you have the developer guide activated:</b>

![guiadesenvolvedor](./guide/guia-desenvolvedor.png)


* <b>2. Open the VisualBasic editor:</b>

![visualbasic](./guide/visual-basic.png)


* <b>3. Import the file makefolder.bas:</b>

![importarquivo](./guide/importar-arquivo.png)


* <b>4. Select cells with folder names:</b>

![selecaopastas](./guide/selecao-itens.png)


* <b>5. Run the script:</b>

![executarsub](./guide/executar-sub.png)



>**Insert a button on your spreadsheet and assign the macro `makefolder.bas`**
>>**This way you can run the script only clicking the button**



* <b>6. Folders will be created automatically:</b>

![folders](./guide/folders.png)

---

### Author

<table>
  <tr>
    <td align="center"><a href="https://github.com/vitoriape"><img style="border-radius: 50%;" src="https://avatars.githubusercontent.com/u/55922652?v=4" width="100px;" alt=""/><br /><sub><b>Vitória Peçanha</b></sub></a></td> 
</table>
