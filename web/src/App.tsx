import { useState, useEffect } from 'react'
import './App.css'
import { Read } from './features/Read';
import { ReadWrite } from './features/ReadWrite';
import { APICalls } from './features/APICalls';

function App() {
  const [officeReady, setOfficeReady] = useState(false);
  const [_, setStorageValue] = useState<string | null>(null);
  const [feature, setFeature] = useState("readwrite");

  useEffect(() => {
    Office.onReady().then(() => {
      console.log("Office is ready");
      setOfficeReady(true);

      setStorageValue(localStorage.getItem("save_key"));
    });

    return () => {
      console.log("Cleanup");
    }
  }, []);

  if (!officeReady) {
    return <div>Loading...</div>
  }

  return (
    <>
      <select value={feature} onChange={(e) => {
        setFeature(e.target.value);
      }}>
        <option value="readwrite">ReadWrite</option>
        <option value="read">Read</option>
        <option value="apicalls">APICalls</option>
      </select>
      {feature === "readwrite" && <ReadWrite />}
      {feature === "read" && <Read />}
      {feature === "apicalls" && <APICalls />}
    </>
  )
}

export default App
