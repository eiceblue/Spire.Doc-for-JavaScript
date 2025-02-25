<template>
  <span>Click the following button to insert header and footer into a Word document</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Sample.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        const inputImgFileName = "Header.png";
        await wasmModule.FetchFileToVFS(inputImgFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        const inputImgFileName_1 = "Footer.png";
        await wasmModule.FetchFileToVFS(inputImgFileName_1,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create word document
        let document = wasmModule.Document.Create();

        document.LoadFromFile(inputFileName);
        let section = document.Sections.get_Item();

        //insert header and footer
        InsertHeaderAndFooter(section,inputImgFileName,inputImgFileName_1);

        // Define the output file name
        const outputFileName = "HeaderAndFooter.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Close();
        document.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }

      function InsertHeaderAndFooter(section,inputImgFileName,inputImgFileName_1) {
        let header = section.HeadersFooters.Header;
        let footer = section.HeadersFooters.Footer;

        //insert picture and text to header
        let headerParagraph = header.AddParagraph();

        let headerPicture = headerParagraph.AppendPicture({imgFile: inputImgFileName});
        //header text
        let text = headerParagraph.AppendText("Demo of Spire.Doc");
        text.CharacterFormat.FontName = "Arial";
        text.CharacterFormat.FontSize = 10;
        text.CharacterFormat.Italic = true;
        headerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

        //border
        headerParagraph.Format.Borders.Bottom.BorderType = wasmModule.BorderStyle.Single;
        headerParagraph.Format.Borders.Bottom.Space = 0.05;


        //header picture layout - text wrapping
        headerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;

        //header picture layout - position
        headerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
        headerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
        headerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
        headerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Top;

        //insert picture to footer
        let footerParagraph = footer.AddParagraph();

        let footerPicture = footerParagraph.AppendPicture({imgFile: inputImgFileName_1});

        //footer picture layout
        footerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;
        footerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
        footerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
        footerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
        footerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Bottom;

        //insert page number
        footerParagraph.AppendField("page number", wasmModule.FieldType.FieldPage);
        footerParagraph.AppendText(" of ");
        footerParagraph.AppendField("number of pages", wasmModule.FieldType.FieldNumPages);
        footerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

        //border
        footerParagraph.Format.Borders.Top.BorderType = wasmModule.BorderStyle.Single;
        footerParagraph.Format.Borders.Top.Space = 0.05;
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
