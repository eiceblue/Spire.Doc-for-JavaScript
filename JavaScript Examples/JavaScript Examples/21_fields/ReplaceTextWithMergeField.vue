<template>
  <span>Click the following button to replace text with a merge field</span>
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
        let inputFileName = "SampleB_2.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Find the text that will be replaced
        let ts = document.FindString("Test", true, true);
  
        // Get the text range of the found string
        let tr = ts.GetAsOneRange();

        // Access the paragraph that owns the text range
        let par = tr.OwnerParagraph;

         // Get the index of the text range in the paragraph
        let index = par.ChildObjects.IndexOf(tr);

        // Create a new merge field in the document
        let field = wasmModule.MergeField.Create(document);
        field.FieldName = "MergeField";

        // Insert the new merge field at the specific position in the paragraph
        par.ChildObjects.Insert(index, field);

        // Remove the original text
        par.ChildObjects.Remove(tr);

        // Define the output file name
        const outputFileName = "ReplaceTextWithMergeField.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
