<template>
  <span>Click the following button to add RichText content control in a Word document</span>
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
        
        // Add a new section 
        let section = document.AddSection();

        // Add a paragraph
        let paragraph = section.AddParagraph();

        // Append textRange for the paragraph
        let txtRange = paragraph.AppendText("The following example shows how to add RichText content control in a Word document. \n");

        // Append textRange
        txtRange = paragraph.AppendText("Add RichText Content Control:  ");

        // Set the font format
        txtRange.CharacterFormat.Italic = true;

        // Create StructureDocumentTagInline for document
        let sdt = wasmModule.StructureDocumentTagInline.Create(document);

        // Add sdt in paragraph
        paragraph.ChildObjects.Add(sdt);

        // Specify the type
        sdt.SDTProperties.SDTType = wasmModule.SdtType.RichText;

        // Set displaying text
        let text = wasmModule.SdtText.Create(true);
        text.IsMultiline = true;
        sdt.SDTProperties.ControlProperties = text;

        // Crate a TextRange
        let rt = wasmModule.TextRange.Create(document);
        rt.Text = "Welcome to use ";
        rt.CharacterFormat.TextColor = wasmModule.Color.get_Green();
        sdt.SDTContent.ChildObjects.Add(rt);

        rt = wasmModule.TextRange.Create(document);
        rt.Text = "Spire.Doc";
        rt.CharacterFormat.TextColor = wasmModule.Color.get_OrangeRed();
        sdt.SDTContent.ChildObjects.Add(rt);
        
        // Define the output file name
        const outputFileName = "AddRichTextContentControl.docx";

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
