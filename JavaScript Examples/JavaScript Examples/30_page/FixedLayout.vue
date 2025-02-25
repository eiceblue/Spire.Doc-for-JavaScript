<template>
  <span>Click the following button to shows fixed layout(e.g., page, lines) of a Word document</span>
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
        let inputFileName = "Template_Docx_3.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Create a FixedLayoutDocument from the loaded document
        let layoutDoc = wasmModule.FixedLayoutDocument.Create(document);
        
        // Get the first line in the first column of the first page
        let line = layoutDoc.Pages.get_Item(0).Columns.get_Item(0).Lines.get_Item(0);

        // Create a StringBuilder to store the output text
        let stringBuilder = [];
        stringBuilder.push("Line: " + line.Text);

        // Get the paragraph that contains the line and append its text to the StringBuilder
        let para = line.Paragraph;
        stringBuilder.push("Paragraph text: " + para.Text + "\n");

        // Get the text content of the first page
        let pageText = layoutDoc.Pages.get_Item(0).Text;
        stringBuilder.push(pageText + "\n");

        // Iterate through each page in the FixedLayoutDocument
        for (let i = 0; i < layoutDoc.Pages.Count; i++) {
            let page = layoutDoc.Pages.get_Item(i);
            // Get all the lines on the current page
            let lines = page.GetChildEntities(wasmModule.LayoutElementType.Line, true);

            // Append the page index and number of lines to the StringBuilder
            stringBuilder.push("Page " + page.PageIndex + " has " + lines.Count + " lines.\n");
        }

        // Append the lines of the first paragraph to the StringBuilder
        // (except runs and nodes in the header and footer).
        stringBuilder.push("The lines of the first paragraph:\n");
        for (let i = 0; i < layoutDoc.GetLayoutEntitiesOfNode(document.FirstChild.Body.Paragraphs.get_Item(0)).Count; i++) {
            let paragraphLine = layoutDoc.GetLayoutEntitiesOfNode(document.FirstChild.Body.Paragraphs.get_Item(0)).get_Item(i);
            stringBuilder.push(paragraphLine.Text.trim() + "\n");
            const x = paragraphLine.Rectangle.X;
            const y= paragraphLine.Rectangle.Y;
            const width = paragraphLine.Rectangle.Width;
            const height = paragraphLine.Rectangle.Height;
            // Create a string with X, Y, width, and height
            const infoString = `{X: ${x}, Y: ${y}, Width: ${width}, Height: ${height}}`;
            stringBuilder.push(infoString+"\n");
        }
        
        // Define the output file name
        const outputFileName = "FixedLayout.txt";

        // Combine all the found data into a single string
        let content = stringBuilder.join("\n");

        // Save the content to the specified path
        wasmModule.FS.writeFile(outputFileName, content);

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
