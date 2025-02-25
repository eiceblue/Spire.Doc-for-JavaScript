<template>
  <span>Click the following button to lock specified sections of Word document</span>
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

        // Create a new document
        const document = wasmModule.Document.Create();

        // Add new sections
        let s1 = document.AddSection();
        let s2 = document.AddSection();

        // Append some text to section 1 and section 2
        s1.AddParagraph().AppendText("Spire.Doc demo, section 1");
        s2.AddParagraph().AppendText("Spire.Doc demo, section 2");

        // Protect the document with AllowOnlyFormFields protection type
        document.Protect({type: wasmModule.ProtectionType.AllowOnlyFormFields, password: "123"});

        // Unprotect section 2
        s2.ProtectForm = false;

        // Define the output file name
        const outputFileName = "LockSpecifiedSections.docx";

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
