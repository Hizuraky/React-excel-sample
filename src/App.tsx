import {useState, useRef} from 'react';
import logo from './logo.png';
import './App.css';
import ExcelJS from "exceljs"
import html2canvas from "html2canvas";

// excelに出力するデータ
const dummyData = [...Array(5)].map((_, i) => {
  return {
    id: i,
    name: "name" + i,
    age: Math.floor(Math.random() * 20) + 20,
  }
})


function App() {
  const componentRef = useRef<HTMLInputElement>(null);
  const [fileName, setName] = useState("");

  const onClickExport = async() => {
    // 出力コンポーネントの高さ・幅取得
    const elementHeight = componentRef.current!.getBoundingClientRect().height;
    const elementWidth = componentRef.current!.getBoundingClientRect().width;

    // 出力コンポーネントをpng形式のbase64に変換
    const componentElement = document.getElementById("rootComponent");
    const base64Component = await html2canvas(componentElement!).then(canvas => canvas.toDataURL("img/png"));

    // excelのブックとシート定義
    const workbook = new ExcelJS.Workbook();
    workbook.addWorksheet("配列用シート");
    workbook.addWorksheet("画像用シート");
    const arraySheet = workbook.getWorksheet("配列用シート");
    const imgSheet = workbook.getWorksheet("画像用シート");

    // 配列出力
    arraySheet.columns = [
      { header: "ID", key: "id" },
      { header: "名前", key: "name" },
      { header: "誕生日", key: "age" }
    ];
    arraySheet.addRows(dummyData);

    // 画像出力
    const image = workbook.addImage({
      base64: base64Component,
      extension: 'png',
    });
    imgSheet.addImage(image,  {
      tl: { col: 0, row: 0 },
      ext: { width: elementWidth, height: elementHeight }
    });
    

    // excel設定・出力 
    const uint8Array = await workbook.xlsx.writeBuffer();
    const blob = new Blob([uint8Array], { type: "application/octet-binary" });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName + ".xlsx" 
    a.click();
    a.remove();
  }
  
  return (
    // 出力したいコンポーネントにidを付ける
    <div className="App" id="rootComponent" ref={componentRef}>
      <header className="App-header">
        <img src={logo} className="App-logo" alt="logo" />
        <p>
          Edit <code>src/App.tsx</code> and save to reload.
        </p>
        <a
          className="App-link"
          href="https://reactjs.org"
          target="_blank"
          rel="noopener noreferrer"
        >
          Learn React
        </a>
        <input onChange={(e) => setName(e.target.value)}/>
        <button onClick={() => onClickExport()} >Excel export</button>
      </header>
    </div>
  );
}

export default App;
