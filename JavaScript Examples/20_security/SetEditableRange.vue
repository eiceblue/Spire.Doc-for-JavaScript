<template>
  <span>Click the following button to set the editable range for the document</span>
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
        let inputFileName = "SetEditableRange.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Protect whole document
        document.Protect({type: wasmModule.ProtectionType.AllowOnlyReading, password: "password"});
        
        // Create tags for permission start and end
        let start = wasmModule.PermissionStart.Create(document, "testID");
        let end = wasmModule.PermissionEnd.Create(document, "testID");
        
        // Add the start and end tags to allow the first paragraph to be edited
        document.Sections.get(0).Paragraphs.get_Item(0).ChildObjects.Insert(0, start);
        document.Sections.get(0).Paragraphs.get_Item(0).ChildObjects.Add(end);

        // Define the output file name
        const outputFileName = "SetEditableRange_out.docx";

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
