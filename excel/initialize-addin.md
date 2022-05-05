# Initialize your Office Add-in
이 문서의 원본은 [여기](https://docs.microsoft.com/en-us/office/dev/add-ins/develop/initialize-add-in)를 참고한다. 


그러나 Office 추가 기능은 라이브러리가 로드될 때까지 Office JavaScript API를 성공적으로 호출할 수 없습니다. 이 문서에서는 코드에서 라이브러리가 로드되었는지 확인할 수 있는 두 가지 방법을 설명합니다.

* Initialize with Office.onReady().
* Initialize with Office.initialize.



> Office.initialize 대신에 Office.onReady()를 사용할 것을 추천한다. 


## Initialize with Office.onReady()

Office.onReady()는 asynchronous method이고Office.js library가 로드되었는지 확인하는 동안 Promise object를 반환한다. 

라이브러리가 로드될 때, 그것은 Office.HostType enum value(Excel, Word, etc.)로  Office client application과 Office.PlatformType enum value (PC, Mac, OfficeOnline, etc.)로 platform을 설명하는 객체로써 Promise를 resolve한다. 

Office.onReady()가 호출되었을 때 라이브러리가 이미 로드되었다면 즉시 Promise는 Resolve한다. 


Office.onReady()를 호출하는 한가지 방법은 callback method를 그것에 전달하는 것이다. 

```jsx
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

















