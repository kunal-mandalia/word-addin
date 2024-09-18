import { useState, useEffect } from 'react'
import './App.css'
import { Read } from './features/Read';

function App() {
  const [officeReady, setOfficeReady] = useState(false);
  const [_, setStorageValue] = useState<string | null>(null);

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
      <Read />
    </>
  )
}

export default App
