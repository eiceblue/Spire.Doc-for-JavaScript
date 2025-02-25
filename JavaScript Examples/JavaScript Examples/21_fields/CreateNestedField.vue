<template>
  <span>Click the following button to create a nested field in a Word document</span>
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
        
        // Create a new document
        const document = wasmModule.Document.Create();

        // Create a new section
        let section = document.AddSection();
        
        // Create a new paragraph
        let paragraph = section.AddParagraph();

        // Create an IF field
        let ifField = wasmModule.IfField.Create(document);
        ifField.Type = wasmModule.FieldType.FieldIf;
        ifField.Code = "IF ";
        paragraph.Items.Add(ifField);

        // Create the embedded IF field
        let ifField2 = wasmModule.IfField.Create(document);
        ifField2.Type = wasmModule.FieldType.FieldIf;
        ifField2.Code = "IF ";
        paragraph.ChildObjects.Add(ifField2);
        paragraph.Items.Add(ifField2);
        paragraph.AppendText("\"200\" < \"50\"   \"200\" \"50\" ");
        let embeddedEnd = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);
        embeddedEnd.Type = wasmModule.FieldMarkType.FieldEnd;
        paragraph.Items.Add(embeddedEnd);
        ifField2.End = embeddedEnd;

        paragraph.AppendText(" > ");
        paragraph.AppendText("\"100\" ");
        paragraph.AppendText("\"Thanks\" ");
        paragraph.AppendText("\"The minimum order is 100 units\"");
        let end = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);
        end.Type = wasmModule.FieldMarkType.FieldEnd;
        paragraph.Items.Add(end);
        ifField.End = end;

        // Update all fields in the document
        document.IsUpdateFields = true;

        // Define the output file name
        const outputFileName = "CreateNestedField.docx";

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
