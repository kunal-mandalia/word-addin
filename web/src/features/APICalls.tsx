import { useEffect, useRef, useState } from "react";

export function APICalls() {
    const [data, setData] = useState<any[]>([]);
    const [debugOutput] = useState("");
    const socket = useRef<WebSocket>();
    const listenMode = useRef<boolean>(false);
    
    // make a call to get mock data from api
    const restCall = async () => {
        const randomId = Math.floor(Math.random() * 100) + 1;
        const response = await fetch(`https://jsonplaceholder.typicode.com/todos/${randomId}`);
        const data = await response.json();
        setData(prev => [...prev, data]);
    }

    // clear data
    const clearData = () => {
        setData([]);
    }

    // send ws message
    const wsCall = (s?: string) => {
        if (socket.current) {
            const ts = new Date().toISOString();            
            socket.current.send(s || ts);
        }
    }

    const listenModeHandler = () => {
        if (listenMode.current && socket.current) {
            socket.current.send("document interaction detected");
            socket.current.send("processing");
            socket.current.send(`data: "stock price of Contoso increased 10% year to date"`);
        }
    }

    useEffect(() => {
        // create a new websocket connection
        const ws = new WebSocket("wss://echo.websocket.org");
        ws.onopen = () => {
            console.log("WebSocket connected");
            ws.send("websocket connected");

            Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, listenModeHandler, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error('Failed to add handler: ' + result.error.message);
                }
            });

            socket.current = ws;
        }

        ws.onmessage = (e) => {
            console.log("Message received: ", e.data);
            setData(prev => [...prev, `wss: ${e.data}`]);
        }



        return () => {
            ws.close();

            Office.context.document.removeHandlerAsync(Office.EventType.DocumentSelectionChanged, { handler: listenModeHandler }, function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    console.error('Failed to remove handler: ' + result.error.message);
                }
            });
        }

    }, []);

    return (
        <div>
            <h2>API Calls</h2>

            <button onClick={() => {
                restCall();
            }}>Rest Call</button>
            {" "}
            <button onClick={() => {
                wsCall();
            }}>WebSocket Call</button>
            {" "}
            <button onClick={() => {
                listenMode.current = !listenMode.current;
                wsCall(`listen mode: ${listenMode.current}`);
            }}>{listenMode.current === true ? "Listen mode: on" : "Listen mode: off"}</button>
            {" "}
            <button onClick={clearData}>Clear</button>
            {" "}
            <div>
                <pre>
                    {data.map((d, i) => {
                        return <div key={i}>{JSON.stringify(d)}</div>
                    })}
                </pre>
            </div>

            <div style={{ color: "red", borderRadius: "8px", padding: "1px", textAlign: "center" }}>
                <pre>
                    {debugOutput}
                </pre>
            </div>
        </div>
    )
}