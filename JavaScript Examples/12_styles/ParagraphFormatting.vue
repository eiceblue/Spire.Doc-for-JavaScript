<template>
  <span>The example demonstrates how to set paragraph formatting options.</span>
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
        let para = sec.AddParagraph();
        para.AppendText("Paragraph Formatting");
        para.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Title});

        para = sec.AddParagraph();
        para.AppendText("This paragraph is surrounded with borders.");
        para.Format.Borders.BorderType = wasmModule.BorderStyle.Single;
        para.Format.Borders.Color = wasmModule.Color.get_Red();

        para = sec.AddParagraph();
        para.AppendText("The alignment of this paragraph is Left.");
        para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;

        para = sec.AddParagraph();
        para.AppendText("The alignment of this paragraph is Center.");
        para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

        para = sec.AddParagraph();
        para.AppendText("The alignment of this paragraph is Right.");
        para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

        para = sec.AddParagraph();
        para.AppendText("The alignment of this paragraph is justified.");
        para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Justify;

        para = sec.AddParagraph();
        para.AppendText("The alignment of this paragraph is distributed.");
        para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Distribute;

        para = sec.AddParagraph();
        para.AppendText("This paragraph has the gray shadow.");
        para.Format.BackColor = wasmModule.Color.get_Gray();

        para = sec.AddParagraph();
        para.AppendText("This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.");
        para.Format.SetLeftIndent(10);
        para.Format.SetRightIndent(10);
        para.Format.SetFirstLineIndent(15);

        para = sec.AddParagraph();
        para.AppendText("The hanging indentation of this paragraph is 15pt.");
        //Negative value represents hanging indentation
        para.Format.SetFirstLineIndent(-15);

        para = sec.AddParagraph();
        para.AppendText("This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.");
        para.Format.AfterSpacing = 20;
        para.Format.BeforeSpacing = 10;
        para.Format.LineSpacingRule = wasmModule.LineSpacingRule.AtLeast;
        para.Format.LineSpacing = 10;

        // Define the output file name
        const outputFileName = "ParagraphFormatting-result.docx";

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
