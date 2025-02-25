<template>
  <span>Click the following button to remove content controls from a Word document</span>
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
        let inputFileName = "RemoveContentControls.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        //Loop through sections
        for (let s = 0; s < document.Sections.Count; s++) {
            let section = document.Sections.get(s);
            for (let i = 0; i < section.Body.ChildObjects.Count; i++) {
                //Loop through contents in paragraph
                if (section.Body.ChildObjects.get(i) instanceof wasmModule.Paragraph) {
                    let para = section.Body.ChildObjects.get(i);
                    for (let j = 0; j < para.ChildObjects.Count; j++) {
                        //Find the StructureDocumentTagInline
                        if (para.ChildObjects.get(j) instanceof wasmModule.StructureDocumentTagInline) {
                            let sdt = para.ChildObjects.get(j);
                            //Remove the content control from paragraph
                            para.ChildObjects.Remove(sdt);
                            j--;
                        }
                    }
                }
                if (section.Body.ChildObjects.get(i) instanceof wasmModule.StructureDocumentTag) {
                    let sdt = section.Body.ChildObjects.get(i);
                    section.Body.ChildObjects.Remove(sdt);
                    i--;
                }
            }
        }
        
        // Define the output file name
        const outputFileName = "RemoveContentControls_out.docx";

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
