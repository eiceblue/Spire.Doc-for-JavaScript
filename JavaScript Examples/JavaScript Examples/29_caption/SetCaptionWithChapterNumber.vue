<template>
  <span>Click the following button to set the caption with the chapter number</span>
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
        let inputFileName = "SetCaptionWithChapterNumber.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        //Get the first section
        let section = document.Sections.get(0);
        
        // Define the caption name for the pictures
        let name = "Caption ";

        // Loop through each paragraph in the section body
        for (let i = 0; i < section.Body.Paragraphs.Count; i++) {
            // Loop through each child object in the current paragraph
            for (let j = 0; j < section.Body.Paragraphs.get_Item(i).ChildObjects.Count; j++) {
                // Check if the child object is a picture
                if (section.Body.Paragraphs.get_Item(i).ChildObjects.get(j) instanceof wasmModule.DocPicture) {
                    let pic1 = section.Body.Paragraphs.get_Item(i).ChildObjects.get(j); 
                    let body = pic1.OwnerParagraph.Owner; 
                    
                    if (body != null) {
                        // Get the index of the paragraph containing the picture
                        let imageIndex = body.ChildObjects.IndexOf(pic1.OwnerParagraph);

                        // Create a new paragraph for the caption
                        let para = wasmModule.Paragraph.Create(document);
                        
                        // Set the caption label
                        para.AppendText(name); 

                        // Add a field for chapter reference
                        let field1 = para.AppendField("test", wasmModule.FieldType.FieldStyleRef);
                        
                        // Set the code for the chapter number
                        field1.Code = " STYLEREF 1 \\s "; 

                        // Append a delimiter between chapter number and caption
                        para.AppendText(" - "); 

                        // Add a field for the picture sequence number
                        let field2 = para.AppendField(name, wasmModule.FieldType.FieldSequence);
                        field2.CaptionName = name; 
                        field2.NumberFormat = wasmModule.CaptionNumberingFormat.Number; 

                        // Insert the new caption paragraph after the picture's paragraph
                        body.Paragraphs.Insert(imageIndex + 1, para);
                      }
                  }
              }
          }
        
        // Update all fields in the document
        document.IsUpdateFields = true;
        
        // Define the output file name
        const outputFileName = "SetCaptionWithChapterNumber.docx";

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
