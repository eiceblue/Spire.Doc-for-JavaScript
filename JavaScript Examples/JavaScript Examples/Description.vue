<template>
  <div class="app-description">
    <h1>Welcome to the Vue Demo</h1>
    <p>
      We are excited to introduce our Vue application, showcasing the powerful capabilities of Spire.Doc for JavaScript in handling Word files. 
      This project seamlessly integrates the functionalities of Spire.Doc for JavaScript through modern web application frameworks. 
    </p>
    <p class="title">Getting Start</p>
    <ol>
      <li>Copy the two files (<span class="getstart">Spire.Doc.Base.js & Spire.Doc.Base.wasm</span>) from <span class="getstart">lib</span> directory to the <span class="getstart">public</span> directory in Vue project.</li>
      <li>Please download the <b>ARIALUNI.TTF</b> font and upload it to the <span class="getstart">public/static/font</span> folder in your project directory. </li>
        <p><b>Note:</b>The ARIALUNI.TTF font is protected by copyright and is only to be used legally for personal and commercial purposes.</p>If your document uses additional fonts, you can also upload these fonts to the public/static/font directory. After uploading, ensure that the font names in your code are adjusted accordingly.
      <li>Expand Spire.Doc from the left menu, select a detailed feature and click the "Start" button to execute.</li>
    </ol>
    <p class="title">Product Introduction</p>
    <p>
      Spire.Doc for JavaScript enables users to create, read, write, convert, and compare Word documents with optimized speed and accuracy. It is a completely independent JavaScript Word class library that doesn't require Microsoft Office installed on systems. 
      With Spire.Doc for JavaScript, developers can efficiently handle Word programming tasks—whether generating new documents from templates, updating existing documents, or performing file conversions.
    </p>
    <p class="title">Technical Architecture</p>
    <p>
      Built using WebAssembly technology, Spire.Doc for JavaScript is fully compatible with popular frameworks such as Vue, React, Angular, and plain JavaScript, 
      enabling developers to easily create and distribute their JavaScript applications on the web.
    </p>
    <p class="title">File Support</p>
    <p>
      Spire.Doc for JavaScript supports Word 97-2003/2007/2010/2013/2016/2019, and it has the capability of converting them to other common formats, like XML, RTF, TXT, EMF, HTML, TIFF, ODT, Markdown, and vice versa. Moreover, it supports converting Word Doc/Docx to PDF, images (PNG, JPEG), PostScript, OFD, XPS, EPUB, PCL (Printer Command Language), and RTF to PDF/HTML, HTML to PDF/Image, Markdown to PDF in high quality.
    </p>
    <p class="title">Contact Us</p>
    <ul>
      <li>For any technical questions regarding the product, feel free to contact us: <a href="mailto:support@e-iceblue.com">support@e-iceblue.com</a>.</li>
      <li>For any questions about purchasing the product, please reach out to: <a href="mailto:sales@e-iceblue.com">sales@e-iceblue.com</a>.</li>
      <li>If you can't find the example you need, contact us for a free custom demo: <a href="mailto:support@e-iceblue.com">support@e-iceblue.com</a>.</li>
    </ul>
  </div>
</template>

<script>
export default {
  name: 'AppDescription',
};
</script>

<style scoped>
.app-description {
  padding: 20px;
  font-family: Arial, sans-serif;
  text-align: left;
}
.title {
  text-align: left;
  font-weight: bold;
  color: #2980B9;
  font-size: 17px;
}

h1 {
  color: #1F618D;
}

p {
  line-height: 1.6;
}
.getstart
{
  color: #E67E22 ;
}
ul {
  list-style-type: none;
  padding: 0;
}
ol {
  list-style-type: style type num;;
  padding: 0;
}
b
{
  font-weight: bold
};
ul li{
  margin: 10px 0;
}
</style>