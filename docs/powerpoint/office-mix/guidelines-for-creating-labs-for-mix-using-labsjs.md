
# <a name="guidelines-for-creating-labs-for-mix-using-labsjs"></a>LabsJS を使用して Mix 用ラボを作成するためのガイドライン



LabsJS ライブラリ (labs.js) は、Office Mix と統合する特殊な Office アドイン (ラボといいます) の作成をサポートしています。Office Mix は、Microsoft PowerPoint を使用してラボを表示します。これらのコンポーネントを「ラボ」と呼ぶ一方で、私たちが作成しているものが特別な Office アドイン (Office Mix アドイン) であることを明確にしておきましょう。

ガイドラインと使用例が示す LabsJS コンテンツは、labs.js JavaScript API の実装に役立ちます。このライブラリは、 [JavaScript API for Office](../../../reference/javascript-api-for-office.md) (Office.js) 上に作成され、Office Mix に組み込まれた アドイン向けに最適化された抽象層が設けられています。


## <a name="general-guidelines"></a>一般的なガイドライン


LabJS API を使用して アドインを作成する際に役立つ一般的なガイドラインは次のとおりです。


### <a name="scripts"></a>スクリプト

labs.js ライブラリは office.js の抽象層であるため、office.js に依存しています。office.js と labs.js の両方のライブラリ ファイルが開発プロジェクトに含まれている必要があります。 

office.js ライブラリは、`<script src="https://sforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>` で参照できます。

labs.js ライブラリは LabsJS SDK に付属しています。また、CDN (<code>https://az592748.vo.msecnd.net/sdk/LabsJS-1.0.4/labs-1.0.4.js</code>) でも labs.js ライブラリを参照することができます。運用バージョンのラボは CDN に格納されているバージョンを参照している必要があることに注意してください。


 >**メモ**: JavaScript ファイル (labs-1.0.4.js) に加えて、ラボ API の TypeScript 定義ファイル (labs-1.0.4.d.ts) が提供されています。この定義ファイルは、TypeScript バージョン 0.9.1.1 に対して作成されました。


### <a name="callbacks-and-error-handling"></a>コールバックとエラー処理

labs.js API のいくつかのメソッドは非同期で処理を実行します。それらの処理のために、この API では  **ILabCallback** という標準コールバック インターフェイスが採用されています。 


```js
function(err, result) {
}
```

コールバック メソッドは、 _err_ と _result_ という 2 つのパラメーターを取ります。[ _err_] フィールドはエラーが発生しない限り  **null** のままです。[ _result_] フィールドは処理の結果を返します。

コールバック処理は、結果がすぐに出る場合でもすぐには発生せず、 別の JavaScript イベント ループが実行されると ( **setTimeout** 呼び出しを通して) 発生します。このコールバック定義を採用すると、選択した promise API と labs.js を簡単に統合できます。たとえば、次の例に示すように、単純な変換メソッドによって、jQuery promises をこれらのコールバックの代わりに使用することができます。




```js
function createCallback<T>(deferred: JQueryDeferred<T>): Labs.Core.ILabCallback<T> {
    return (err, data) => {
        if (err) {
            deferred.reject(err);
        }
        else {
            deferred.resolve(data);
        }
    };
}
```


### <a name="lab-host-and-defaultlabhost"></a>ラボのホストと DefaultLabHost

ラボのホスト ( **ILabHost**) は、ラボの開発をサポートする基になるドライバーです。既定では、これは office.js と統合するホストに設定されます。

テストのためや labhost.html 内でラボを実行するためには、シミュレーション環境で稼働するホストへの切り替えが必要になります。次のコード例は、クエリ パラメーターを使用してこれを実行する方法を示しています。または、ラボの アドイン が別のプラットフォームと統合するように  **DefaultHostBuilder** を変更することもできます。




```js
Labs.DefaultHostBuilder = function () {
    if (window.location.href.indexOf("PostMessageLabHost") !== -1) {
        return new Labs.PostMessageLabHost("test", parent, "*");
    } else {
        return new Labs.OfficeJSLabHost();
    }
};
```


### <a name="initialization"></a>初期化

初期化すると、ラボとそのホスト間の通信経路が確立されます。以下を呼び出してラボを初期化してください。


```js
Labs.connect((err, connectionResponse) => {});
```

初期化したら、labs.js API の他のメソッドを呼び出します。 _connectionResponse_ パラメーターには、ホスト、ユーザーに関する情報、およびその他の接続に関連した情報が含まれています。返される値について詳しくは「 [Labs.Core.IConnectionResponse](../../../reference/office-mix/labs.core.iconnectionresponse.md)」を参照してください。


### <a name="time-format"></a>時刻の形式

Labs.js は、1970 年 1 月 1 日 (UTC) からのミリ秒単位の経過時間の数値を格納します。これは、JavaScript [Date オブジェクト](http://msdn.microsoft.com/en-us/library/ie/cd9w2te4%28v=vs.94%29.aspx)の日付形式と一致します。


### <a name="timeline"></a>タイムライン

ラボでは、レッスン プレーヤーのタイムラインも操作できます。タイムラインを使用すると、ラボはレッスン プレーヤーに次のスライドへ進むように指示できます。タイムライン オブジェクトは、 **Labs.getTimeline** メソッドを呼び出して取得します。


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="handling-events"></a>イベントを処理する


LabsJS のイベント API では、ラボ固有のイベントを追跡し、イベントに応答したり対応したりできるようにするためにイベント ハンドラーを追加することが可能です。イベント メソッドは 3 つあり ( **ModeChanged**、 **Activate**、および  **Deactivate**)、それらは  **EventTypes** オブジェクトにあります。 


### <a name="mode-change"></a>モードの変更

**ModeChanged** イベントは、指定したラボが編集モードから表示モードに変更されたときに発生します。編集モードは、ラボが PowerPoint の編集モードで表示されているときに表示されます。表示モードは、PowerPoint がスライド ショーを表示しているとき、またはラボが Office Mix のレッスン プレーヤーで表示されているときに表示されます。表示モードでは、常にユーザーがラボを使用しているときに見えるものが表示されます。ラボの構成は、編集モードで行うことができます。

コールバックに渡される  **ModeChangedEventData** オブジェクトのデータには、現在のモードに関する情報が含まれています。次のコードは、 **ModeChanged** イベントの使用方法を示しています。




```js
Labs.on(Labs.Core.EventTypes.ModeChanged, (data) => {
    var modeChangedEvent = <Labs.Core.ModeChangedEventData> data;
    this.switchToMode(modeChangedEvent.mode);
});
```


### <a name="activate"></a>アクティブ化

**activate** イベントは、ラボがある PowerPoint のスライドがレッスン プレーヤーでアクティブになったときに発生します。


```js
Labs.on(Labs.Core.EventTypes.Activate, (data) => {
    //  is now on the active slide
});
```


### <a name="deactivate"></a>非アクティブ化

**deactivate** イベントは、ラボがある PowerPoint のスライドがアクティブなスライドではなくなったときに発生します。


```js
Labs.on(Labs.Core.EventTypes.Deactivate, (data) => {                
    //  is no longer on the active slide
});
```


### <a name="timeline"></a>タイムライン

ラボでは、レッスン プレーヤーのタイムラインも操作できます。タイムラインを使用すると、ラボはレッスン プレーヤーに、次のスライドに進むように伝えます。タイムライン オブジェクトは、 **Labs.getTimeline** メソッドを呼び出して取得します。


```js
Labs.getTimeline().next({}, (err, unused) => { });
```


## <a name="additional-resources"></a>その他のリソース



- [Office Mix アドイン](../../powerpoint/office-mix/office-mix-add-ins.md)
    
