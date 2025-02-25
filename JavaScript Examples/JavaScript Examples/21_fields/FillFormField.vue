<template>
  <span>Click the following button to fill form field</span>
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
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "FillFormField.doc";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        const inputFileName1 = "User.xml";
        await wasmModule.FetchFileToVFS(inputFileName1,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Read XML content from the input file
        const data = wasmModule.FS.readFile(inputFileName1);
        
        // Create a TextDecoder instance to decode the Uint8Array into a string using UTF-8 encoding
        const decoder = new TextDecoder('utf-8');
        
        // Decode the Uint8Array data into a string
        const stringData = decoder.decode(data);

        // Parse the decoded string as XML using DOMParser
        const xmlDocument = new DOMParser().parseFromString(stringData,'application/xml');

        // Select the first <user> element from the parsed XML document
        let user = xmlDocument.querySelector("user");
        
        // Fill data into form fields
        for (let i = 0; i < document.Sections.get(0).Body.FormFields.Count; i++) {
            let field = document.Sections.get(0).Body.FormFields.get_Item(i);
            let path = field.Name ;
            let propertyNode = user.querySelector(path);
            
            // Check if the property node exists
            if (propertyNode != null) {
              // Switch based on the type of the form field
              switch (field.Type) {
                  case wasmModule.FieldType.FieldFormTextInput:
                      // If the field is a text input, set its text to the property node's text content
                      field.Text = propertyNode.textContent;
                      break;

                  case wasmModule.FieldType.FieldFormDropDown:
                      // If the field is a dropdown, find the correct item to select
                      let combox = field;
                      for (let i = 0; i < combox.DropDownItems.Count; i++) {
                        // Check if the item text matches the property value
                          if (combox.DropDownItems.get_Item(i).Text === propertyNode.Value) {
                             // Set the selected index
                              combox.DropDownSelectedIndex = i;
                              break;
                          }
                          // Special case for the "country" field to select "Others" if applicable
                          if (field.Name == "country" && combox.DropDownItems.get_Item(i).Text === "Others") {
                              combox.DropDownSelectedIndex = i;
                          }
                      }
                      break;

                    case wasmModule.FieldType.FieldFormCheckBox:
                       // If the field is a checkbox, check if it should be checked
                      if (propertyNode.textContent) {
                          let checkBox = field;
                          // Set the checkbox to checked
                          checkBox.Checked = true;
                      }
                      break;
                }
              }   
          }

        // Define the output file name
        const outputFileName = "FillFormField_out.doc";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Doc});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/msword"});

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
