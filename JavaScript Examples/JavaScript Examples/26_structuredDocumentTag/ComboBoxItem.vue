<template>
  <span>Click the following button to add, select and remove combo box item</span>
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
        let inputFileName = "ComboBox.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        //Get the combo box from the file
        for (let i = 0; i < document.Sections.Count; i++) {
            let section = document.Sections.get_Item(i);
            for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
                let bodyObj = section.Body.ChildObjects.get(j);
                if (bodyObj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTag) {
                    //If SDTType is ComboBox
                    if (bodyObj.SDTProperties.SDTType === wasmModule.SdtType.ComboBox) {
                        let combo = bodyObj.SDTProperties.ControlProperties;
                        //Remove the second list item
                        combo.ListItems.RemoveAt(1);
                        //Add a new item
                        let item = wasmModule.SdtListItem.Create("D", "D");
                        combo.ListItems.Add(item);

                        //If the value of list items is "D"
                        for (let i = 0; i < combo.ListItems.Count; i++) {
                            let sdtItem = combo.ListItems.get_Item(i);
                            if (sdtItem.Value === "D") {
                                //Select the item
                                combo.ListItems.SelectedValue = sdtItem;
                            }
                        }
                    }
                }
            }
        }
        
        // Define the output file name
        const outputFileName = "ComboBoxItem.docx";

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
