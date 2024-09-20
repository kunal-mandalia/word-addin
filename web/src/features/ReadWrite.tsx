export function ReadWrite() {

    async function insertCitation(replace?: boolean) {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const space = replace ? "" : " ";
            const r0 = selection.insertText(`${space}The quick brown fox `, replace ? Word.InsertLocation.replace : Word.InsertLocation.end);
            const r1 = r0.insertText("jumped", Word.InsertLocation.end);

            const c1 = r1.insertComment(`"jumped" - Open AI: openai.com`);
            c1.contentRange.hyperlink = "https://www.openai.com";

            const c2 = r1.insertComment(`"jumped" - Perplexity AI: perplexity.ai`);
            c2.contentRange.hyperlink = "https://www.perplexity.ai";

            const r2 = r1.insertText(" over the lazy ", Word.InsertLocation.end);
            const r3 = r2.insertText("dog", Word.InsertLocation.end);
            const c3 = r3.insertComment("dog - Wikipedia: wikipedia.org");
            c3.contentRange.hyperlink = "https://www.wikipedia.org";
            await context.sync();
        });
    }

    return (
        <div>
            <h2>ReadWrite</h2>
            <h3>Citation</h3>
            <button onClick={() => {
                insertCitation(false);
            }}>Append</button>{" "}
            <button onClick={() => { insertCitation(true) }}>Replace</button>
        </div>
    )
}