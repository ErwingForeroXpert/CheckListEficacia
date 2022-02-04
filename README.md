<div id="top"></div>

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]

<br />
<div align="center">
  <a href="https://github.com/ErwingForeroXpert/CheckListEficacia">
    <img src="images/screenshot.png" alt="Logo">
  </a>

  <h3 align="center">Automatización Checklist Eficacia</h3>

  <p align="center">
    <br />
    <a href="https://github.com/othneildrew/Best-README-Template"><strong>Explore the docs »</strong></a>
    <br />
  </p>
</div>

<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#comenzar">Comenzar</a>
      <ul>
        <li><a href="#prerequisitos">Prerequisitos</a></li>
        <li><a href="#instalacion">Instalacion</a></li>
        <li><a href="#generar ejecutable">Ejecutable</a></li>
      </ul>
    </li>
    <li><a href="#contacto">Contacto</a></li>
  </ol>
</details>

## Comenzar

### Prerequisitos

* [Python](https://www.python.org/downloads/) <@3.8>
* Microsoft Excel (Recomendable 2019)
* [VS code](https://code.visualstudio.com/) (opcional)

### Instalación

_Para la instalación del proyecto solo es necesario:_

1. Clonar el repositorio
   ```sh
   git clone https://github.com/ErwingForeroXpert/CheckListEficacia.git
   ```
2. Instalar `pipenv`, omitir si ya esta instalado
   ```sh
   pip install pipenv
   ```
3. Instalar las librerias necesarias
   ```sh
     pipenv install
   ```
4. Ejecutar el proyecto
   ```sh
     python index.py
   ```

<p align="right">(<a href="#top">back to top</a>)</p>

### Generar Ejecutable

_Para Generar el archivo ejecutable es necesario lo siguiente:_

1. Verificar que el entorno virtual esta activo
   ```sh
   pipenv --venv
   ```
   Deberia mostar la ruta del entorno virtual activo.
   
2. Ejecutar el siguiente comando:
   ```sh
   pipenv run build
   ```
3. Crear el archivo ".env" en el directorio `./dist/`, el cual debe contener lo siguiente:
   ```sh
     URL_EFICACIA="URL"
     USER="USUARIO"
     PASSWORD="CONTRASEÑA"
    ```
   

## Licencia

Distributed under the MIT License. See `LICENSE.txt` for more information.

<p align="right">(<a href="#top">back to top</a>)</p>

## Contacto

Erwing FC  - erwing.forero@xpertgroup.co

Project Link: [https://github.com/ErwingForeroXpert/CheckListEficacia](https://github.com/ErwingForeroXpert/CheckListEficacia)

<p align="right">(<a href="#top">back to top</a>)</p>

<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/ErwingForeroXpert/CheckListEficacia
[contributors-url]: https://github.com/ErwingForeroXpert/CheckListEficacia/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/ErwingForeroXpert/CheckListEficacia
[forks-url]: https://github.com/ErwingForeroXpert/CheckListEficacia/network/members
[stars-shield]: https://img.shields.io/github/stars/ErwingForeroXpert/CheckListEficacia
[stars-url]: https://github.com/ErwingForeroXpert/CheckListEficacia/stargazers
[issues-shield]: https://img.shields.io/github/issues/ErwingForeroXpert/CheckListEficacia
[issues-url]: https://github.com/ErwingForeroXpert/CheckListEficacia/issues
[license-shield]: https://img.shields.io/github/license/ErwingForeroXpert/CheckListEficacia
[license-url]: https://github.com/ErwingForeroXpert/CheckListEficacia/blob/develop/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-Linkedin-blue
[linkedin-url]: https://www.linkedin.com/in/erwing-forero-castro-586781133
[product-screenshot]: images/screenshot.png
