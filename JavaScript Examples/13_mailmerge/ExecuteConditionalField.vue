<template>
  <span>The example shows how to execute conditional field including merge field.</span>
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

        let doc = wasmModule.Document.Create();
        //Add a new section
        let section = doc.AddSection();
        //Add a new paragraph for a section
        let paragraph = section.AddParagraph();

        CreateIFField1(doc, paragraph);
        paragraph = section.AddParagraph();
        CreateIFField2(doc, paragraph);

        let fieldName = ["Count", "Age"];
        let fieldValue = ["2", "30"];

        doc.MailMerge.Execute(fieldName, fieldValue);
        doc.IsUpdateFields = true;

        // Define the output file name
        const outputFileName = "ExecuteConditionalField-result.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
        
        // Clean up resources
        doc.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
      function CreateIFField1( document,  paragraph) {
        let ifField = wasmModule.IfField.Create(document);
        ifField.Type = wasmModule.FieldType.FieldIf;
        ifField.Code = "IF ";
        paragraph.Items.Add(ifField);

        paragraph.AppendField("Count", wasmModule.FieldType.FieldMergeField);
        paragraph.AppendText(" > ");
        paragraph.AppendText("\"1\" ");
        paragraph.AppendText("\"Greater than one\" ");
        paragraph.AppendText("\"Less than one\"");

        let end = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);
        end.Type = wasmModule.FieldMarkType.FieldEnd;
        paragraph.Items.Add(end);

        ifField.End = end;
      }

    function CreateIFField2( document,  paragraph) {
        let ifField = wasmModule.IfField.Create(document);
        ifField.Type = wasmModule.FieldType.FieldIf;
        ifField.Code = "IF ";
        paragraph.Items.Add(ifField);

        paragraph.AppendField("Age", wasmModule.FieldType.FieldMergeField);
        paragraph.AppendText(" > ");
        paragraph.AppendText("\"50\" ");
        paragraph.AppendText("\"The old man\" ");
        paragraph.AppendText("\"The young man\"");

        let end = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);
        end.Type = wasmModule.FieldMarkType.FieldEnd;
        paragraph.Items.Add(end);

        ifField.End = end;
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
