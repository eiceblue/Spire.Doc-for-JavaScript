<template>
    <span>Click the following button to extract text from textboxes in Word document</span>
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
                await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

                // Load the sample file into the virtual file system (VFS)
                let inputFileName = "ExtractTextFromTextBoxes.docx";
                await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

                // Create a new document
                const document = wasmModule.Document.Create();

                // Load a document from the virtual file system
                document.LoadFromFile(inputFileName);

                let sw = [];

                // Verify whether the document contains a textbox or not.
                if (document.TextBoxes.Count > 0) {
                    // Traverse the document.
                    for (let i = 0; i < document.Sections.Count; i++) {
                        let section = document.Sections.get(i);
                        for (let j = 0; j < section.Paragraphs.Count; j++) {
                            let p = section.Paragraphs.get_Item(j);
                            for (let k = 0; k < p.ChildObjects.Count; k++) {
                                let obj = p.ChildObjects.get(k);
                                if (obj.DocumentObjectType == wasmModule.DocumentObjectType.TextBox) {
                                    let textbox = obj;
                                    for (let a = 0; a < textbox.ChildObjects.Count; a++) {
                                        let objt = textbox.ChildObjects.get(a);
                                        //Extract text from paragraph in TextBox.
                                        if (objt.DocumentObjectType == wasmModule.DocumentObjectType.Paragraph) {
                                            sw.push(objt.Text);
                                        }

                                        // Extract text from Table in TextBox.
                                        if (objt.DocumentObjectType == wasmModule.DocumentObjectType.Table) {
                                            let table = objt;
                                            ExtractTextFromTables(table, sw);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // Combine all the found data into a single string
                let content = sw.join("\n");

                // Define the output file name
                const outputFileName = "ExtractTextFromTextBoxes.txt";

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

        const ExtractTextFromTables = (table, sw) => {
            for (let i = 0; i < table.Rows.Count; i++) {
                let row = table.Rows.get(i);
                for (let j = 0; j < row.Cells.Count; j++) {
                    let cell = row.Cells.get(j);
                    for (let k = 0; k < cell.Paragraphs.Count; k++) {
                        let paragraph = cell.Paragraphs.get_Item(k);
                        sw.push(paragraph.Text);
                    }
                }
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
