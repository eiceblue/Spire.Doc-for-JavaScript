<template>
  <span>Click the following button to insert OLE as the icon to the document via the stream</span>
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
        let inputFileName = "example.zip";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        const imageFileName = "example.png";
        await wasmModule.FetchFileToVFS(imageFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        
        // Create a new document
        const document = wasmModule.Document.Create();
        
        // Add a section
        let sec = document.AddSection();

        // Add a paragraph
        let par = sec.AddParagraph();

        // Create an OLE stream from the specified input file
        let stream = wasmModule.Stream.CreateByFile(inputFileName);

        // Load the image into a DocPicture object
        let picture = wasmModule.DocPicture.Create(document);
        picture.LoadImage(imageFileName);

        // Insert the OLE object using the created stream and picture
        let obj = par.AppendOleObject({
            oleStream: stream,       
            olePicture: picture,      
            fileExtension: "zip"      
        });

        // Display the OLE object as an icon instead of the content
        obj.DisplayAsIcon = true;   
        
        // Define the output file name
        const outputFileName = "InsertOLEAsIconViaStream.docx";

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
