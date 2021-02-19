# Excel2Asciidoc
VBA macro that translate Excel to Asciidoc

<!-- PROJECT LOGO -->
<br />
<p align="center">
  <h3 align="center">Excel2Asciidoc</h3>

  <p align="center">
    VBA macro that translate Excel to Asciidoc
  </p>
</p>

<!-- TABLE OF CONTENTS -->
<details open="open">
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#about-the-project">About The Project</a> </li>
    <li><a href="#installation-and-usage">Installation and Usage</a></li>
    <li><a href="#license">License</a></li>
  </ol>
</details>

<!-- ABOUT THE PROJECT -->
## About The Project

![Product Name Screen Shot][product-screenshot]

Excel is very useful calc software. 
Sometime, Excel is used to write a technical document, but that is not well received.
Of course, if you write as document, it is better to use a word processor, word.

Now, we use Excel for writing technical document. Excel is good for making tables, and editing pictures. Especially, document, sentence.. table.. picture.. sentence.. picture. At that case, we think excel is best.

In the future, not only asciidoc but others markdown or textile etc.

This is only sample code, and there are many limitation.

## Installation and Usage

### 1. Make new excel file.
### 2. Open VBE.
* Alt+F11 will open the VBE
### 3. Import files(*.bas, *.cls).
* import files from  test/src
### 4. Once save as xlsm file.
* Save file as xlsm file. etc. excel2ascii.xlsm.
### 5. Make sample sheet.
* Alt + F8, and execute MakeSampleSheet.
* You get the sheet below.
![samplesheet][samplesheet]
### 6. Convert Excel to Asciidoc.
* Alt + F8, and execute MakeDocumentAndPic.
* You get a sheet1.adoc and a $B$31.png  in the folder excel2ascii.xlsm exists.
* you can view the sheet1.adoc at some asciidoctor viewer.


<!-- USAGE EXAMPLES
## Usage

Use this space to show useful examples of how a project can be used. Additional screenshots, code examples and demos work well in this space. You may also link to more resources.

_For more examples, please refer to the [Documentation](https://example.com)_
 -->

<!-- LICENSE -->
## License

Distributed under the MIT License. See `LICENSE` for more information.

<!-- CONTACT
## Contact

Your Name - [@your_twitter](https://twitter.com/your_username) - email@example.com

Project Link: [https://github.com/toramameseven/Excel2Asciidoc](https://github.com/toramameseven/Excel2Asciidoc)
 -->

<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->

[product-screenshot]: images/screenshot.png
[samplesheet]: images/SampleSheet.png