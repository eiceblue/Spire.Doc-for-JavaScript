<template>
  <span>The example demonstrates how to set all available character formatting options.</span>
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

        let document = wasmModule.Document.Create();
        let sec = document.AddSection();
        let titleParagraph = sec.AddParagraph();
        titleParagraph.AppendText("Font Styles and Effects ");
        titleParagraph.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Title});

        let paragraph = sec.AddParagraph();
        let tr = paragraph.AppendText("Strikethough Text");
        tr.CharacterFormat.IsStrikeout = true;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Shadow Text");
        tr.CharacterFormat.IsShadow = true;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Small caps Text");
        tr.CharacterFormat.IsSmallCaps = true;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Double Strikethough Text");
        tr.CharacterFormat.DoubleStrike = true;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Outline Text");
        tr.CharacterFormat.IsOutLine = true;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("AllCaps Text");
        tr.CharacterFormat.AllCaps = true;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Text");
        tr = paragraph.AppendText("SubScript");
        tr.CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SubScript;

        tr = paragraph.AppendText("And");
        tr = paragraph.AppendText("SuperScript");
        tr.CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SuperScript;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Emboss Text");
        tr.CharacterFormat.Emboss = true;
        tr.CharacterFormat.TextColor = wasmModule.Color.get_White();

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Hidden:");
        tr = paragraph.AppendText("Hidden Text");
        tr.CharacterFormat.Hidden = true;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Engrave Text");
        tr.CharacterFormat.Engrave = true;
        tr.CharacterFormat.TextColor = wasmModule.Color.get_White();

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("WesternFonts中文字体");
        tr.CharacterFormat.FontNameAscii = "Calibri";
        tr.CharacterFormat.FontNameNonFarEast = "Calibri";
        tr.CharacterFormat.FontNameFarEast = "Simsun";

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Font Size");
        tr.CharacterFormat.FontSize = 20;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Font Color");
        tr.CharacterFormat.TextColor = wasmModule.Color.get_Red();

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Bold Italic Text");
        tr.CharacterFormat.Bold = true;
        tr.CharacterFormat.Italic = true;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Underline Style");
        tr.CharacterFormat.UnderlineStyle = wasmModule.UnderlineStyle.Single;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Highlight Text");
        tr.CharacterFormat.HighlightColor = wasmModule.Color.get_Yellow();

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Text has shading");
        tr.CharacterFormat.TextBackgroundColor = wasmModule.Color.get_Green();

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Border Around Text");
        tr.CharacterFormat.Border.BorderType = wasmModule.BorderStyle.Single;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Text Scale");
        tr.CharacterFormat.TextScale = 150;

        paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
        tr = paragraph.AppendText("Character Spacing is 2 point");
        tr.CharacterFormat.CharacterSpacing = 2;

        // Define the output file name
        const outputFileName = "CharacterFormatting-result.docx";

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
