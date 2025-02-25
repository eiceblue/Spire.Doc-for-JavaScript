<template>
  <span>The example shows how to get text based on style name in a Word document.</span>
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
        let inputFileName = "GetTextByStyleName.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Load document from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Create string builder
        let builder = []

        //Loop through sections
        for (let i = 0; i < doc.Sections.Count; i++) {
            let section = doc.Sections.get(i);
            //Loop through paragraphs
            for (let j = 0; j < section.Paragraphs.Count; j++) {
                let para = section.Paragraphs.get_Item(j);
                //Find the paragraph whose style name is "Heading1"
                if (para.StyleName == "Heading1") {
                    //Write the text of paragraph
                    builder.push(para.Text + "\n");
                }
            }
        }

        // Define the output file name
        const outputFileName = "GetTextByStyleName-result.txt";
        //Write the contents in a TXT file
        wasmModule.FS.writeFile(outputFileName, builder.join("\n"));
        doc.Close();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type: "text/plain"});

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
