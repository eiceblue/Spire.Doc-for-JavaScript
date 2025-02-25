<template>
  <span>The example demonstrates how to create styles and assign them to paragraph.</span>
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

        //Initialize a document
        let document = wasmModule.Document.Create();
        let sec = document.AddSection();

        //Add default title style to document and modify
        let titleStyle = document.AddStyle(wasmModule.BuiltinStyle.Title);

        titleStyle.CharacterFormat.FontName = "cambria";
        titleStyle.CharacterFormat.FontSize = 28;

        titleStyle.CharacterFormat.TextColor = wasmModule.Color.FromArgb(42, 123, 136);

        //judge if it is Paragraph Style and then set paragraph format
        if (titleStyle instanceof wasmModule.ParagraphStyle) {
            let ps = titleStyle;
            ps.ParagraphFormat.Borders.Bottom.BorderType = wasmModule.BorderStyle.Single;
            ps.ParagraphFormat.Borders.Bottom.Color = wasmModule.Color.FromArgb(42, 123, 136);
            ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5;
            ps.ParagraphFormat.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;
        }
        //Add default normal style and modify
        let normalStyle = document.AddStyle(wasmModule.BuiltinStyle.Normal);

        normalStyle.CharacterFormat.FontName = "cambria";
        normalStyle.CharacterFormat.FontSize = 11;

        //Add default heading1 style
        let heading1Style = document.AddStyle(wasmModule.BuiltinStyle.Heading1);

        heading1Style.CharacterFormat.FontName = "cambria";
        heading1Style.CharacterFormat.FontSize = 14;

        heading1Style.CharacterFormat.Bold = true;
        heading1Style.CharacterFormat.TextColor = wasmModule.Color.FromArgb(42, 123, 136);
        //Add default heading2 style
        let heading2Style = document.AddStyle(wasmModule.BuiltinStyle.Heading2);

        heading2Style.CharacterFormat.FontName = "cambria";
        heading2Style.CharacterFormat.FontSize = 12;

        heading2Style.CharacterFormat.Bold = true;

        //List style
        let bulletList = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);

        bulletList.CharacterFormat.FontName = "cambria";
        bulletList.CharacterFormat.FontSize = 12;

        bulletList.Name = "bulletList";
        document.ListStyles.Add(bulletList);
        //Apply the style
        let paragraph = sec.AddParagraph();
        paragraph.AppendText("Your Name");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Title});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Address, City, ST ZIP Code | Telephone | Email");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Normal});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Objective");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading1});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("To get started right away, just click any placeholder text (such as this) and start typing to replace it with your own.");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Normal});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Education");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading1});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("DEGREE | DATE EARNED | SCHOOL");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Major:Text");
        paragraph.ListFormat.ApplyStyle("bulletList");
        paragraph = sec.AddParagraph();
        paragraph.AppendText("Minor:Text");
        paragraph.ListFormat.ApplyStyle("bulletList");
        paragraph = sec.AddParagraph();
        paragraph.AppendText("Related coursework:Text");
        paragraph.ListFormat.ApplyStyle("bulletList");

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Skills & Abilities");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading1});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("MANAGEMENT");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Think a document that looks this good has to be difficult to format? Think again! To easily apply any text formatting you see in this document with just a click, on the Home tab of the ribbon, check out Styles.");
        paragraph.ListFormat.ApplyStyle("bulletList");

        paragraph = sec.AddParagraph();
        paragraph.AppendText("COMMUNICATION");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("You delivered that big presentation to rave reviews. Don’t be shy about it now! This is the place to show how well you work and play with others.");
        paragraph.ListFormat.ApplyStyle("bulletList");

        paragraph = sec.AddParagraph();
        paragraph.AppendText("LEADERSHIP");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Are you president of your fraternity, head of the condo board, or a team lead for your favorite charity? You’re a natural leader—tell it like it is!");
        paragraph.ListFormat.ApplyStyle("bulletList");

        paragraph = sec.AddParagraph();
        paragraph.AppendText("Experience");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading1});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("JOB TITLE | COMPANY | DATES FROM - TO");
        paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

        paragraph = sec.AddParagraph();
        paragraph.AppendText("This is the place for a brief summary of your key responsibilities and most stellar accomplishments.");
        paragraph.ListFormat.ApplyStyle("bulletList");

        // Define the output file name
        const outputFileName = "Styles-result.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
        
        // Clean up resources
        document.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
