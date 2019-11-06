const fs = require('fs');
const docx = require("docx");
const { Document, Media, Packer, Paragraph, Table, TableRow, TableCell, TableHeader, TextRun, WidthType, VerticalAlign, AlignmentType, TextWrappingType, TextWrappingSide, HeadingLevel, TableOfContents } = docx;

const doc = new Document();

let json = fs.readFileSync('./testData/41476.json');

let obj = JSON.parse(json);

let paragraphs = [];
paragraphs.push(new TableOfContents("Contents", {
    hyperlink: true,
    headingStyleRange: "1-2",
    //stylesWithLevels: [new StyleLevel("MySpectacularStyle", 1)],
}))
let HEADERNAMES = [
    {
        obj: "description",
        value: "Description"
    },
    {
        obj: "manufacturer",
        value: "Manufacturer"
    },
    {
        obj: "number",
        value: "Drawing Number"
    },
    {
        obj: "referencecount",
        value: "Quantity"
    }
]

function recurseObj(obj, level) {
    const { children, data } = obj;

    //console.log(`Level:${level}: ${obj.data.description}, style: Heading ${level}`)
    paragraphs.push(new Paragraph({
        text: obj.data.description,
        heading: HeadingLevel[`HEADING_${level}`]
    }))

    let datakeys = Object.keys(data);
    datakeys.forEach(key => {
        paragraphs.push(new Paragraph(`${key}: ${data[key]}`))
    })
    if (children && Array.isArray(children)) {
        let tableCont = [];
        let filtered = children.filter(element => element.data.filetype !== "CWR");
        if (filtered && filtered.length > 0) {
            tableCont.push(new TableRow({
                tableHeader: true,
                children: HEADERNAMES.map(head => {
                    return new TableCell({
                        verticalAlign: VerticalAlign.CENTER,
                        children: [
                            new Paragraph({
                                alignment: AlignmentType.CENTER,
                                text: head.value,
                                style: "Normal"
                            })
                        ]
                    })
                })
            }))
            filtered
                .forEach(element => {
                    let cells = [];
                    HEADERNAMES.forEach(name => {
                        let textValue = element.data[name.obj];
                        if (!isNaN(textValue)){textValue = textValue.toString()}
                            cells.push(new TableCell({
                                verticalAlign: VerticalAlign.CENTER,
                                children: [
                                    new Paragraph({
                                        alignment: AlignmentType.CENTER,
                                        text: textValue,
                                        style: "Normal"
                                    })
                                ]
                            }))
                    })
                    tableCont.push(new TableRow({
                        children: cells
                    }))

                });
            paragraphs.push(new Table({
                rows: tableCont,
                width: {
                    type: WidthType.PERCENTAGE,
                    size: 100
                }
            }));
        }
    }

    if (children && Array.isArray(children)) {
        level++;
        //console.log(level)
        children.forEach(element => {
            if (element.data.filetype !== "CWR") {
                recurseObj(element, level)
            }
        });
    }
    //return obj2;
};
recurseObj(obj, 1);

doc.addSection({ children: paragraphs });

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});














//--------Test
function addHeaders(names) {
    return new TableRow({
        tableHeader: true,
        children: names.map(head => {
            return new TableCell({
                verticalAlign: VerticalAlign.CENTER,
                children: [
                    new Paragraph({
                        alignment: AlignmentType.CENTER,
                        text: head,
                        style: "Normal"
                    })
                ]
            })

        })
    })
}


let headers = addHeaders(['Test', 'Headers', 'Print']);
let image = Media.addImage(doc, fs.readFileSync("./img/40018-small.jpg"), 200, 200, {})
let row = new TableRow({
    children: [new TableCell({
        margins: {
            top: 20,
            bottom: 20,
            right: 20,
            left: 20
        },
        children: [new Paragraph(image)]
    })]
})
let table = new Table({ rows: [headers, row], width: { type: WidthType.PERCENTAGE, size: 100 } });