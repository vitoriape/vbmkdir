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

Índice
=================
<!--ts-->
   * [Sobre](#sobre)
   * [Ferramentas](#ferramentas)
   * [Referências](#referências)
   * [Instalação](#instalação)
   * [Autor](#autor)
 
---
   
### Sobre

Este projeto é um script feito em [VBA](https://docs.microsoft.com/pt-br/office/vba/library-reference/concepts/getting-started-with-vba-in-office) que cria pastas automaticamente a partir de células selecionadas no Excel.


### Ferramentas

Development of this script utilizes the tools listed below:

- [Git](https://git-scm.com/)
- [Excel](https://support.microsoft.com/en-us/excel)
- [Visual Basic for Applications](https://docs.microsoft.com/pt-br/office/vba/api/overview/excel)

                  
### Referências

Para mais informações sobre o uso da instrução `Do (...) Loop` e da `MkDir`, além da função `Dir` no Visual Basic for Applications, leia a documentação da Microsoft:

- [Instrução Do...Loop](https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/doloop-statement)
- [Instrução MkDir](https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/mkdir-statement)
- [Função Dir](https://docs.microsoft.com/pt-br/office/vba/language/reference/user-interface-help/dir-function)

---

### Instalação

```bash
# Clone esse repositório
$ git clone <https://github.com/vitoriape/vbmkdir>
```

* <b>1. Verifique se você tem a guia desenvolvedor ativada:</b>

![guiadesenvolvedor](./guide/guia-desenvolvedor.png)


* <b>2. Abra o editor do VisualBasic:</b>

![visualbasic](./guide/visual-basic.png)


* <b>3. Importe o arquivo makefolder.bas:</b>

![importarquivo](./guide/importar-arquivo.png)


* <b>4. Selecione as células com os nomes das pastas:</b>

![selecaopastas](./guide/selecao-itens.png)


* <b>5. Rode o script:</b>

![executarsub](./guide/executar-sub.png)



>**Crie um botão na sua planilha e atribua a macro deste projeto `makefolder.bas`**
>>**Dessa forma você pode rodar o script apenas clicando no botão**



* <b>6. As pastas serão criadas automaticamente:</b>

![folders](./guide/folders.png)

---

### Autor

<table>
  <tr>
    <td align="center"><a href="https://github.com/vitoriape"><img style="border-radius: 50%;" src="https://avatars.githubusercontent.com/u/55922652?v=4" width="100px;" alt=""/><br /><sub><b>Vitória Peçanha</b></sub></a></td> 
</table>
