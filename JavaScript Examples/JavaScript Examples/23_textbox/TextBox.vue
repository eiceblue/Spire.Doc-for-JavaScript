<template>
  <span>Click the following button to insert a textbox into Word</span>
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

        let section = document.AddSection();

        InsertTextbox(section);

        // Define the output file name
        const outputFileName = "Textbox.docx";

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
    const InsertTextbox = (section) => {
      let paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
      paragraph = section.AddParagraph();
      paragraph = section.AddParagraph();

      // Insert and format the first textbox.
      let textBox1 = paragraph.AppendTextBox(240, 35);
      textBox1.Format.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
      textBox1.Format.LineColor = wasmModule.Color.get_Gray();
      textBox1.Format.LineStyle = wasmModule.TextBoxLineStyle.Simple;
      textBox1.Format.FillColor = wasmModule.Color.get_DarkSeaGreen();
      let para = textBox1.Body.AddParagraph();
      let txtrg = para.AppendText("Textbox 1 in the document");
      txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
      txtrg.CharacterFormat.FontSize = 14;
      txtrg.CharacterFormat.TextColor = wasmModule.Color.get_White();
      para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

      // Insert and format the second textbox.
      paragraph = section.AddParagraph();
      paragraph = section.AddParagraph();
      paragraph = section.AddParagraph();
      paragraph = section.AddParagraph();
      let textBox2 = paragraph.AppendTextBox(240, 35);
      textBox2.Format.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
      textBox2.Format.LineColor = wasmModule.Color.get_Tomato();
      textBox2.Format.LineStyle = wasmModule.TextBoxLineStyle.ThinThick;
      textBox2.Format.FillColor = wasmModule.Color.get_Blue();
      textBox2.Format.LineDashing = wasmModule.LineDashing.Dot;
      para = textBox2.Body.AddParagraph();
      txtrg = para.AppendText("Textbox 2 in the document");
      txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
      txtrg.CharacterFormat.FontSize = 14;
      txtrg.CharacterFormat.TextColor = wasmModule.Color.get_Pink();
      para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

      // Insert and format the third textbox.
      paragraph = section.AddParagraph();
      paragraph = section.AddParagraph();
      paragraph = section.AddParagraph();
      paragraph = section.AddParagraph();
      let textBox3 = paragraph.AppendTextBox(240, 35);
      textBox3.Format.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
      textBox3.Format.LineColor = wasmModule.Color.get_Violet();
      textBox3.Format.LineStyle = wasmModule.TextBoxLineStyle.Triple;
      textBox3.Format.FillColor = wasmModule.Color.get_Pink();
      textBox3.Format.LineDashing = wasmModule.LineDashing.DashDotDot;
      para = textBox3.Body.AddParagraph();
      txtrg = para.AppendText("Textbox 3 in the document");
      txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
      txtrg.CharacterFormat.FontSize = 14;
      txtrg.CharacterFormat.TextColor = wasmModule.Color.get_Tomato();
      para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
    }

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
