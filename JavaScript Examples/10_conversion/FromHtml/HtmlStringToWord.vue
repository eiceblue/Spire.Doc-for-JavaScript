<template>
  <span>The example demonstrates how to convert HTML string to Word.</span>
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
        let inputFileName = "HtmlStringToWord.txt";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Get html string.
        let HTML = "<html><head><style type=\"text/css\">li,p{font-family:'Lucida Sans Unicode';font-size:14pt;}</style></head><body><font size=\"16pt\" color=\"blue\"><h2 align=\"center\">Spire.Doc</h2></font><p><b>Edition:</b></p>";
          HTML+="<ul type=\"disc\"><li><span style='color:green;'>Free Edition</span></li><li>Trial version</li><li><span style='color:red'>A month free for trial version</span></li></ul></ul><p><b>Platform:</b></p><ul type=\"square\">";
          HTML+="<li>.NET</li><li>WPF</li><li>Silverlight</li></ul><table border=\"1\" width=\"90%\"><tr><th>Main Functions of Spire.Doc</th></tr><tr><td>Convert File Documents with High Quality</td></tr> <tr><td>Richest Word Document Features Support</td></tr>";
          HTML+="<tr><td>Simple & Easy to Process Pre-Existing Word Documents</td></tr><tr><td>Other Technical Features</td></tr></table></body></html>";

        // Create a new document
        const document = wasmModule.Document.Create();

        // Create a new section
        let section = document.AddSection();

        // Create a new paragraph
        let paragraph = section.AddParagraph();

        //Append html string.
        paragraph.AppendHTML(HTML.toString('utf8',0,HTML.length));

        // Define the output file name
        const outputFileName = "HtmlStringToWord-result.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        document.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
