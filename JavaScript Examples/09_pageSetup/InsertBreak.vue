<template>
  <span>Click the following button to insert break in a Word document.</span>
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
        function SetPage(section) {
          //the unit of all measures below is point, 1point = 0.3528 mm
          section.PageSetup.PageSize = wasmModule.PageSize.A4();
          section.PageSetup.Margins.Top = 72;
          section.PageSetup.Margins.Bottom = 72;
          section.PageSetup.Margins.Left = 89.85;
          section.PageSetup.Margins.Right = 89.85;
        }

        function InsertCover(section) {
          let small = wasmModule.ParagraphStyle.Create(section.Document);
          small.Name = "small";
          small.CharacterFormat.FontName = "Arial";
          small.CharacterFormat.FontSize = 9;
          small.CharacterFormat.TextColor = wasmModule.Color.get_Gray();
          section.Document.Styles.Add(small);


          let paragraph = section.AddParagraph();
          paragraph.AppendText("The sample demonstrates how to insert section break.");
          paragraph.ApplyStyle(small.Name);

          let title = section.AddParagraph();
          let text = title.AppendText("Field Types Supported by Spire.Doc");
          text.CharacterFormat.FontName = "Arial";
          text.CharacterFormat.FontSize = 20;
          text.CharacterFormat.Bold = true;
          title.Format.BeforeSpacing
            = section.PageSetup.PageSize.Height / 2 - 3 * section.PageSetup.Margins.Top;
          title.Format.AfterSpacing = 8;
          title.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

          paragraph = section.AddParagraph();
          paragraph.AppendText("e-iceblue Spire.Doc team.");
          paragraph.ApplyStyle(small.Name);
          paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;
        }

        function InsertContent(section) {
          let list = wasmModule.ParagraphStyle.Create(section.Document);
          list.Name = "list";
          list.CharacterFormat.FontName = "Arial";
          list.CharacterFormat.FontSize = 11;
          list.ParagraphFormat.LineSpacing = 1.5 * 12;
          list.ParagraphFormat.LineSpacingRule = wasmModule.LineSpacingRule.Multiple;
          section.Document.Styles.Add(list);

          let title = section.AddParagraph();
          let text = title.AppendText("Field type list:");
          title.ApplyStyle(list.Name);

          let first = true;

          for (let type in wasmModule.FieldType) {
            if (wasmModule.FieldType.hasOwnProperty(type)) {
              if (wasmModule.FieldType[type] === wasmModule.FieldType.FieldUnknown
                || wasmModule.FieldType[type] === wasmModule.FieldType.FieldNone || wasmModule.FieldType[type] === wasmModule.FieldType.FieldEmpty) {
                continue;
              }
              let paragraph = section.AddParagraph();
              paragraph.AppendText(type + "is supported in Spire.Doc");

              if (first) {
                paragraph.ListFormat.ApplyNumberedStyle();
                first = false;
              } else {
                paragraph.ListFormat.ContinueListNumbering();
              }
              paragraph.ApplyStyle(list.Name);
            }
          }
        }


        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        //Create word document
        let document = wasmModule.Document.Create();

        //Add section
        let section = document.AddSection();

        //page setup
        SetPage(section);

        //Add cover.
        InsertCover(section);

        //insert a break code
        section = document.AddSection();
        section.AddParagraph().InsertSectionBreak({ breakType: wasmModule.SectionBreakType.NewPage });

        //add content
        InsertContent(section);

        // Define the output file name
        const outputFileName = "InsertBreak_out.docx";

        // Save the document to the specified path
        document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

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
