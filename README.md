# Spire.Doc-for-JavaScript

A powerful Word processing library allowing developers to create, read, edit, and convert Word documents in JavaScript.

[![Foo](https://i.imgur.com/E7SCiVs.png)](https://www.e-iceblue.com/Introduce/doc-for-javascript.html)

[Product Page](https://www.e-iceblue.com/Introduce/doc-for-javascript.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-doc-f6.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)

[Spire.Doc for JavaScript](https://www.e-iceblue.com/Introduce/doc-for-javascript.html) is a powerful API that enables developers to create, read, write, convert, and compare Word documents with high speed and accuracy. It integrates Word document creation into JavaScript applications without needing Microsoft Word installed on either the development or target systems.

Supporting Word formats from 97-2003 to 2019, Spire.Doc for JavaScript can convert DOC/DOCX files to various formats, including XML, RTF, TXT, HTML, ODT, and Markdown. It also supports converting to PDF, images (PNG, JPEG), PostScript, XPS, EPUB, PCL, and more. Additionally, it allows high-quality conversions from RTF to PDF/HTML, HTML to PDF/Image, and Markdown to PDF.

This API offers a fast, reliable, and cost-effective solution for Word document processing and format conversion, streamlining workflows and improving productivity.

### 100% Independent JavaScript API - No Microsoft Office Needed
Spire.Doc for JavaScript is a fully independent JavaScript library for working with Word documents, with no need for Microsoft Office to be installed on your system. It provides an efficient, fast, and scalable solution for generating and processing Word documents, offering a clear advantage over traditional Microsoft Office Automation methods. 

### High-Quality in Convertion 
Spire.Doc for JavaScript offers high-quality conversion between Word Doc/Docx and various formats, including XML, RTF, TXT, EMF, HTML, ODT, Markdown, and more. It also supports converting Word documents to PDF, images (PNG, JPEG), PostScript, OFD, XPS, EPUB, and PCL, as well as converting RTF to PDF/HTML, HTML to PDF/Image, and Markdown to PDF. Additionally, users can save Word documents to streams or web responses.

### Full Support for Advanced Word Document Features
Spire.Doc for JavaScript provides comprehensive support for creating and managing Word documents, offering full access to key document elements such as pages, sections, headers, footers, footnotes, paragraphs, lists, tables, bookmarks, comments, images, and background settings. It also supports advanced features like drawing objects, shapes, textboxes, images, and OLE objects, enabling users to generate dynamic Word documents with rich content and formatting.

### Effortless Processing of Existing Word Documents
Spire.Doc for JavaScript simplifies the process of working with existing Word documents. It supports essential document operations such as search and replace, alignment adjustments, page breaks, field filling, document concatenation, and document copying, enabling efficient and streamlined document management.

### High Performance
● High-quality conversion.
● High processing speed.

### More Technical Supports
Spire.Doc for JavaScript enables developers to create and manipulate Word documents within a 64-bit Node.js environment. It offers a wide range of features for document generation, editing, and conversion, providing flexibility and power for building robust applications.

## Vue Examples

### Convert Word to PDF in JavaScript
```JavaScript
<template>
  <span>The example demonstrates how to convert Word to PDF.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      try {
        wasmModule = window.wasmModule;
        if (wasmModule) {
          // Load font file
          await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

          // Load the Word file
          let inputFileName = "ToPDFTemplate.docx";
          await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

          // Create a new document object
          const document = wasmModule.Document.Create();
          // Load the Word file
          document.LoadFromFile(inputFileName);

          // Define output file name
          const outputFileName = "ToPDF-result.pdf";

          // Save to PDF
          document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF });

          // Clean up resources
          document.Dispose();

          // Read the saved file and convert to a Blob object
          const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
          const modifiedFile = new Blob([modifiedFileArray], { type: "application/pdf" });

          // Set download properties
          downloadName.value = outputFileName;
          downloadUrl.value = URL.createObjectURL(modifiedFile);
        }
      } catch (error) {
        console.error("Error processing the document:", error);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
```

### Convert Word to Text in JavaScript
```JavaScript
<template>
  <span>The example shows how to convert Word document to txt file.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "WordToTxt.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        // Define the output file name
        const outputFileName = "WordToTxt-result.txt";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Txt});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

        // Clean up resources
        doc.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
```

### Count the number of words in a document using JavaScript
```JavaScript
<template>
  <span>Click the following button to count the number of words in a Word document.</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Template_Docx_1.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();
        // Load the template file
        document.LoadFromFile(inputFileName);

        //Count the number of words.
        let content = [];
        content.push("CharCount: " + document.BuiltinDocumentProperties.CharCount + "\n");
        content.push("CharCountWithSpace: " + document.BuiltinDocumentProperties.CharCountWithSpace + "\n");
        content.push("WordCount: " + document.BuiltinDocumentProperties.WordCount + "\n");

        // Define the output file name
        const outputFileName = "CountWordsNumber_out.txt";

        //Save to file.
        FS.writeFile(outputFileName, content.join(""))

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'text/plain' });

        // Clean up resources
        document.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
```
[Product Page](https://www.e-iceblue.com/Introduce/doc-for-javascript.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-doc-f6.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)
