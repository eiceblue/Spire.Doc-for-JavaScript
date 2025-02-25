<template>
  <span>Click the following button to link header and footer in a Word document.</span>
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
        let inputFileName1 = "Template_N1.docx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = "Template_N2.docx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        //Load the source file
        let srcDoc = wasmModule.Document.Create()
        srcDoc.LoadFromFile(inputFileName1);

        //Load the destination file
        let dstDoc = wasmModule.Document.Create();
        dstDoc.LoadFromFile(inputFileName2);

        //Link the headers and footers in the source file
        srcDoc.Sections.get_Item(0).HeadersFooters.Header.LinkToPrevious = true;
        srcDoc.Sections.get_Item(0).HeadersFooters.Footer.LinkToPrevious = true;

        //Clone the sections of source to destination
        for (let i = 0; i < srcDoc.Sections.Count; i++) {
          let section = srcDoc.Sections.get_Item(i);
          dstDoc.Sections.Add(section.Clone());
        }

        // Define the output file name
        const outputFileName = "LinkHeadersFooters_out.docx";

        // Save the document to the specified path
        dstDoc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Clean up resources
        dstDoc.Dispose();
        srcDoc.Dispose();

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
