<template>
  <span>Click the following button to copy header and footer between Word documents</span>
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
        let inputFileName = "HeaderAndFooter.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        const inputFileName_1 = "Template.docx";
        await wasmModule.FetchFileToVFS(inputFileName_1,"",`${import.meta.env.BASE_URL}static/data/`);


        //Load the source file
        let doc1 = wasmModule.Document.Create();
        doc1.LoadFromFile(inputFileName);

        //Get the header section from the source document
        let header = doc1.Sections.get(0).HeadersFooters.Header;

        //Load the destination file
        let doc2 = wasmModule.Document.Create();
        doc2.LoadFromFile(inputFileName_1);

        //Copy each object in the header of source file to destination file
        for (let i = 0; i < doc2.Sections.Count; i++) {
            let section = doc2.Sections.get_Item(i);
            for (let j = 0; j < header.ChildObjects.Count; j++) {
                let obj = header.ChildObjects.get(j);
                section.HeadersFooters.Header.ChildObjects.Add(obj.Clone());
            }
        }

        // Define the output file name
        const outputFileName = "CopyHeaderAndFooter_output.docx";

        // Save the document to the specified path
        doc2.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc1.Close();
        doc2.Close();
        doc1.Dispose();
        doc2.Dispose();

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
