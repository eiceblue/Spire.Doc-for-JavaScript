<template>
  <span>The example shows how to retrieve style names that are applied in a Word document.</span>
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
        let inputFileName = "RetrieveStyle.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load a template document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Traverse all paragraphs in the document and get their style names through StyleName property
        let styleName = "";
        for (let i = 0; i < doc.Sections.Count; i++) {
            let section = doc.Sections.get(i);
            for (let j = 0; j < section.Paragraphs.Count; j++) {
                let paragraph = section.Paragraphs.get_Item(j);
                styleName += paragraph.StyleName + "\r\n";
            }
        }

        // Define the output file name
        const outputFileName = "RetrieveStyle-result.txt";

        //Write the contents in a TXT file
        wasmModule.FS.writeFile(outputFileName, styleName);
        doc.Close();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
