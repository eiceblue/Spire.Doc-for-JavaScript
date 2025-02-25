<template>
  <span>Click the following button to remove hyperlinks from a Word document</span>
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
        let inputFileName = "Hyperlinks.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);        

        //Load Document
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get all hyperlinks
        let hyperlinks = FindAllHyperlinks(doc);

        //Flatten all hyperlinks
        for (let i = hyperlinks.length - 1; i >= 0; i--) {
            FlattenHyperlinks(hyperlinks[i]);
        }

        // Define the output file name
        const outputFileName = "RemoveHyperlinks_output.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc.Close();
        doc.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }

    //Create a method FindAllHyperlinks() to get all the hyperlinks from the sample document
    function FindAllHyperlinks(document) {
    let hyperlinks = [];
    //Iterate through the items in the sections to find all hyperlinks
    for (let i = 0; i < document.Sections.Count; i++) {
        let section = document.Sections.get(i);
        for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
            let sec = section.Body.ChildObjects.get(j);
            if (sec.DocumentObjectType == wasmModule.DocumentObjectType.Paragraph) {
                for (let k = 0; k < sec.ChildObjects.Count; k++) {
                    let para = sec.ChildObjects.get(k);
                    if (para.DocumentObjectType == wasmModule.DocumentObjectType.Field) {
                        let field = para;

                        if (field.Type == wasmModule.FieldType.FieldHyperlink) {
                            hyperlinks.push(field);
                        }
                    }
                }
            }
        }
    }
    return hyperlinks;
}

        // Flatten the hyperlink field
function FlattenHyperlinks(field) {
    let ownerParaIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph);
    let fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field);
    let sepOwnerPara = field.Separator.OwnerParagraph;
    let sepOwnerParaIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph);
    let sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator);
    let endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End);
    let endOwnerParaIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph);

    FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

    field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex);

    for (let i = sepOwnerParaIndex; i >= ownerParaIndex; i--) {
        if (i == sepOwnerParaIndex && i == ownerParaIndex) {
            for (let j = sepIndex; j >= fieldIndex; j--) {
                field.OwnerParagraph.ChildObjects.RemoveAt(j);

            }
        } else if (i == ownerParaIndex) {
            for (let j = field.OwnerParagraph.ChildObjects.Count - 1; j >= fieldIndex; j--) {
                field.OwnerParagraph.ChildObjects.RemoveAt(j);
            }

        } else if (i == sepOwnerParaIndex) {
            for (let j = sepIndex; j >= 0; j--) {
                sepOwnerPara.ChildObjects.RemoveAt(j);
            }
        } else {
            field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i);
        }
    }
}

        //Remove the font color and underline format of the hyperlinks
function FormatFieldResultText(ownerBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex) {
    for (let i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++) {
        let para = ownerBody.ChildObjects.get(i);
        if (i == sepOwnerParaIndex && i == endOwnerParaIndex) {
            for (let j = sepIndex + 1; j < endIndex; j++) {
                FormatText(para.ChildObjects.get(j));
            }

        } else if (i == sepOwnerParaIndex) {
            for (let j = sepIndex + 1; j < para.ChildObjects.Count; j++) {
                FormatText(para.ChildObjects.get(j));
            }
        } else if (i == endOwnerParaIndex) {
            for (let j = 0; j < endIndex; j++) {
                FormatText(para.ChildObjects.get(j));
            }
        } else {
            for (let j = 0; j < para.ChildObjects.Count; j++) {
                FormatText(para.ChildObjects.get(j));
            }
        }
    }
}
function FormatText(tr) {
    //Set the text color to black
    tr.CharacterFormat.TextColor = wasmModule.Color.get_Black();
    //Set the text underline style to none
    tr.CharacterFormat.UnderlineStyle = wasmModule.UnderlineStyle.None;
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
