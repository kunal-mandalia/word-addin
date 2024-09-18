import { useEffect, useState } from "react"

export function Read() {
    const [title, setTitle] = useState("");
    const [paragraphs, setParagraphs] = useState<Word.Paragraph[]>([]);

    useEffect(() => {
        // get the title of Word document
        Office.context.document.getFilePropertiesAsync((result) => {
            setTitle(result.value.url);
        });

        Word.run(async (context) => {
            const paragraphs = context.document.getSelection().paragraphs;
            paragraphs.load();
            await context.sync();
            setParagraphs(paragraphs.items);
        });
    }, []);

    return (
        <div>
            <h2>Read</h2>
            <h3>Title</h3>
            <p>Document name: {title}</p>
            
            <h3>Paragraphs</h3>
            {paragraphs.map((paragraph, index) => (
                <p key={index}>{paragraph.text}</p>
            ))}
        </div>
    )
}