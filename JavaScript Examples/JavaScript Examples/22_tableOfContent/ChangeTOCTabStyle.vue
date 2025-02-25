<template>
  <span>Click the following button to change tab style of TOC in a Word document</span>
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
        let inputFileName = "Template_Toc.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Loop through sections
        for (let i = 0; i < document.Sections.Count; i++) {
            let section = document.Sections.get_Item(i);
            // Loop through content of section
            for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
                let obj = section.Body.ChildObjects.get(j);
                // Find the structure document tag
                if (obj instanceof wasmModule.StructureDocumentTag) {
                    let tag = obj;
                    // Find the paragraph where the TOC1 locates
                    for (let k = 0; k < tag.ChildObjects.Count; k++) {
                        let cObj = tag.ChildObjects.get(k);
                        if (cObj instanceof wasmModule.Paragraph) {
                            let para = cObj;
                            if (para.StyleName == "TOC2") {
                                // Set the tab style of paragraph
                                for (let a = 0; a < para.Format.Tabs.Count; a++) {
                                    let tab = para.Format.Tabs.get_Item(a);
                                    tab.Position = tab.Position + 20;
                                    tab.TabLeader = wasmModule.TabLeader.NoLeader;
                                }
                            }
                        }
                    }
                }
            }
        }

        // Define the output file name
        const outputFileName = "ChangeTOCTabStyle.docx";

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
