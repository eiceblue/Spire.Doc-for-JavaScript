<template>
  <span>Click the following button to add content controls in a Word document</span>
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

        // Load the sample image into the virtual file system (VFS)
        let inputFileName = "logo.png";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Add a new section
        let section = document.AddSection();
        
        // Add a paragraph
        let paragraph = section.AddParagraph();

        // Append textRange in the paragraph
        let txtRange = paragraph.AppendText("The following example shows how to add content controls in a Word document.");
        paragraph = section.AddParagraph();

        // Add Combo Box Content Control
        paragraph = section.AddParagraph();
        txtRange = paragraph.AppendText("Add Combo Box Content Control:  ");
        txtRange.CharacterFormat.Italic = true;
        let sd = wasmModule.StructureDocumentTagInline.Create(document);
        paragraph.ChildObjects.Add(sd);
        sd.SDTProperties.SDTType = wasmModule.SdtType.ComboBox;
        let cb = wasmModule.SdtComboBox.Create();
        cb.ListItems.Add(wasmModule.SdtListItem.Create("Spire.Doc"));
        cb.ListItems.Add(wasmModule.SdtListItem.Create("Spire.XLS"));
        cb.ListItems.Add(wasmModule.SdtListItem.Create("Spire.PDF"));
        sd.SDTProperties.ControlProperties = cb;
        let rt = wasmModule.TextRange.Create(document);
        rt.Text = cb.ListItems.get_Item(0).DisplayText;
        sd.SDTContent.ChildObjects.Add(rt);

        section.AddParagraph();

        // Add Text Content Control
        paragraph = section.AddParagraph();
        txtRange = paragraph.AppendText("Add Text Content Control:  ");
        txtRange.CharacterFormat.Italic = true;
        sd = wasmModule.StructureDocumentTagInline.Create(document);
        paragraph.ChildObjects.Add(sd);
        sd.SDTProperties.SDTType = wasmModule.SdtType.Text;
        let text = wasmModule.SdtText.Create(true);
        text.IsMultiline = true;
        sd.SDTProperties.ControlProperties = text;
        rt = wasmModule.TextRange.Create(document);
        rt.Text = "Text";
        sd.SDTContent.ChildObjects.Add(rt);

        section.AddParagraph();

        // Add Picture Content Control
        paragraph = section.AddParagraph();
        txtRange = paragraph.AppendText("Add Picture Content Control:  ");
        txtRange.CharacterFormat.Italic = true;
        sd = wasmModule.StructureDocumentTagInline.Create(document);
        paragraph.ChildObjects.Add(sd);
        sd.SDTProperties.SDTType = wasmModule.SdtType.Picture;
        let pic = wasmModule.DocPicture.Create(document);
        pic.Width = 10;
        pic.Height = 10;

        pic.LoadImage(inputFileName);
        sd.SDTContent.ChildObjects.Add(pic);

        section.AddParagraph();

        // Add Date Picker Content Control
        paragraph = section.AddParagraph();
        txtRange = paragraph.AppendText("Add Date Picker Content Control:  ");
        txtRange.CharacterFormat.Italic = true;
        sd = wasmModule.StructureDocumentTagInline.Create(document);
        paragraph.ChildObjects.Add(sd);
        sd.SDTProperties.SDTType = wasmModule.SdtType.DatePicker;
        let date = wasmModule.SdtDate.Create();
        date.CalendarType = wasmModule.CalendarType.Default;
        date.DateFormat = "yyyy.MM.dd";
        date.FullDate = wasmModule.DateTime.get_Now();
        sd.SDTProperties.ControlProperties = date;
        rt = wasmModule.TextRange.Create(document);
        rt.Text = "1990.02.08";
        sd.SDTContent.ChildObjects.Add(rt);

        section.AddParagraph();

        // Add Drop-Down List Content Control
        paragraph = section.AddParagraph();
        txtRange = paragraph.AppendText("Add Drop-Down List Content Control:  ");
        txtRange.CharacterFormat.Italic = true;
        sd = wasmModule.StructureDocumentTagInline.Create(document);
        paragraph.ChildObjects.Add(sd);
        sd.SDTProperties.SDTType = wasmModule.SdtType.DropDownList;
        let sddl = wasmModule.SdtDropDownList.Create();
        sddl.ListItems.Add(wasmModule.SdtListItem.Create("Harry"));
        sddl.ListItems.Add(wasmModule.SdtListItem.Create("Jerry"));
        sd.SDTProperties.ControlProperties = sddl;
        rt = wasmModule.TextRange.Create(document);
        rt.Text = sddl.ListItems.get_Item(0).DisplayText;
        sd.SDTContent.ChildObjects.Add(rt);
        
        // Define the output file name
        const outputFileName = "AddContentControls.docx";

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
