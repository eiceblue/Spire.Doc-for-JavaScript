<template>
  <span>Click the following button to create IF field in a Word document</span>
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

        // Define a method of creating an IF Field
        CreateIfField(document, paragraph);

        // Define merged data
        let fieldName = ["Count"];
        let fieldValue = ["2"];

        // Merge data into the IF Field
        document.MailMerge.Execute({fieldNames: fieldName, fieldValues: fieldValue});

        // Update all fields in the document
        document.IsUpdateFields = true;

        // Define the output file name
        const outputFileName = "CreateIFField.docx";

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
    const CreateIfField = (document, paragraph) => {
      // Create an IF field in the document
      let ifField = wasmModule.IfField.Create(document);
      ifField.Type = wasmModule.FieldType.FieldIf;
      ifField.Code = "IF ";
      
      // Add the IF field to the paragraph
      paragraph.Items.Add(ifField);

      // Append a merge field named "Count" to the paragraph
      paragraph.AppendField("Count", wasmModule.FieldType.FieldMergeField);

      // Append text to the paragraph to complete the IF condition
      paragraph.AppendText(" > ");
      paragraph.AppendText("\"100\" ");
      paragraph.AppendText("\"Thanks\" ");
      paragraph.AppendText("\"The minimum order is 100 units\"");

      // Create an end marker for the IF field
      let end = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);

      // Add the end marker to the paragraph
      end.Type = wasmModule.FieldMarkType.FieldEnd;
      paragraph.Items.Add(end);

      // Link the end marker to the IF field
      ifField.End = end;
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
