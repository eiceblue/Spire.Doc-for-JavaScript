<template>
  <span>Click the following button to update CheckBox content control</span>
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
        let inputFileName = "CheckBoxContentControl.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Call tagInlines
        let tagInlines = GetAllTags(document);

        // Get the controls
        for (let i = 0; i < tagInlines.length; i++) {
            // Get the type
            let type = tagInlines[i].SDTProperties.SDTType;

            // Update the status
            if (type === wasmModule.SdtType.CheckBox) {
                let scb = tagInlines[i].SDTProperties.ControlProperties;
                if (scb.Checked) {
                    scb.Checked = false;
                } else {
                    scb.Checked = true;
                }
            }
          }
        
        // Define the output file name
        const outputFileName = "UpdateCheckBox.docx";

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
    const GetAllTags = (document) => {
      let tagInlines = [];
      // Travel document sections
      for (let i = 0; i < document.Sections.Count; i++) {
          let section = document.Sections.get(i);
          for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
              let obj = section.Body.ChildObjects.get(j);
              // Travel document paragraphs
              if (obj.DocumentObjectType === wasmModule.DocumentObjectType.Paragraph) {
                  for (let k = 0; k < obj.ChildObjects.Count; k++) {
                      let pobj = obj.ChildObjects.get(k);
                      // Get StructureDocumentTagInline
                      if (pobj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTagInline) {
                          tagInlines.push(pobj);
                      }
                  }
              }
          }
        }
        return tagInlines;
    }
    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
