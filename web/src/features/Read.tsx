import { useEffect, useState } from "react"

export function Read() {
    const [title, setTitle] = useState("");
    const [selectedText, setSelectedText] = useState("");
    const [selectedParagraphText, setSelectedParagraphText] = useState("");
    
    const [paragraphs, setParagraphs] = useState<Word.Paragraph[]>([]);

    async function syncSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error('Action failed. Error: ' + asyncResult.error.message);
            } else {
                console.log('Selected data: ' + asyncResult.value);
                setSelectedText(asyncResult.value as string);
            }
        });

        Word.run(async (context) => {
            const selectedParagraphs = context.document.getSelection().paragraphs;
            selectedParagraphs.load("items");

            await context.sync();
            setSelectedParagraphText(selectedParagraphs.items.map(p => p.text).join("\n"));
        });
    }

    async function getContent() {
        // get the title of Word document
        Office.context.document.getFilePropertiesAsync((result) => {
            const parsed = result.value.url.split("/");
            setTitle(parsed[parsed.length - 1]);
        });

        Word.run(async (context) => {
            const paragraphs = context.document.body.paragraphs;
            paragraphs.load("items");

            await context.sync();
            setParagraphs(paragraphs.items);
        });
    }

    useEffect(() => {
        async function init() {
            await Word.run(async (context) => {
                const document = context.document;
                document.onParagraphChanged.add(getContent);
            });

            Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, syncSelection, function(result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error('Failed to add handler: ' + result.error.message);
                }
            });
            getContent();
        }
        init();
    }, []);

    return (
        <div>
        <div>
            <h2>Read</h2>
            <h3>Title</h3>
            <p>{title}</p>

            <h3>Content</h3>
            <pre style={{ maxHeight: 400, overflow: "scroll" }}>
                {JSON.stringify(paragraphs, null, 4)}
            </pre>

            <h3>Context</h3>

            <div>Paragraph: </div>
            <p>{selectedParagraphText}</p>

            <div>Selection: </div>
            <p>{selectedText}</p>
        </div>

            <div>
                <button onClick={getContent}>Sync</button>
            </div>
        </div>
    )
}