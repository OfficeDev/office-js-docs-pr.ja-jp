
# <a name="persisting-add-in-state-and-settings"></a>アドインの状態および設定を保持する

Office アドイン は、基本的にブラウザー コントロールのステートレス環境で動作する Web アプリケーションです。したがって、ご使用のアドインでは、そのアドインを使用するセッション間で特定の操作または機能を継続して維持するためのデータを保持することが必要な場合があります。たとえば、アドインには、ユーザーの優先ビューや既定の場所など、アドインで保存しておき、アドインが次回初期化されたときにリロードする必要があるカスタム設定または他の値が含まれる場合があります。

その場合は、次の方法を使用します。


- アドインの種類によって異なる場所に格納されたプロパティ バッグに、名前/値のペアとしてデータを格納する JavaScript API for Office のメンバーを使用します。
    
- 基になるブラウザー コントロールによって提供される技術である、ブラウザーの Cookie、または HTML5 Web ストレージ ([localStorage](http://msdn.microsoft.com/en-us/library/cc848902%28v=vs.85%29.aspx) または [sessionStorage](http://msdn.microsoft.com/en-us/library/cc197020%28v=vs.85%29.aspx)) を使用します。
    
この記事では、アドインの状態を保持する JavaScript API for Office の使い方について説明します。ブラウザーの Cookie および Web ストレージの使用例については、「 [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a>JavaScript API for Office を使用してアドインの状態および設定を保持する


JavaScript API for Office には、次の表に示すように、セッション間でアドインの状態を保存するために [Settings](http://dev.office.com/reference/add-ins/shared/settings) オブジェクト、 [RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) オブジェクト、および [CustomProperties](http://dev.office.com/reference/add-ins/outlook/CustomProperties) オブジェクトが用意されています。すべてのケースで、保存された設定値は、それを作成したアドインの [Id](http://msdn.microsoft.com/en-us/library/67c4344a-935c-09d6-1282-55ee61a2838b%28Office.15%29.aspx) にのみ関連付けられます。



|**オブジェクト**|**アドインの種類のサポート**|**ストレージの場所**|**サポートされる Office のホスト**|
|:-----|:-----|:-----|:-----|
|[設定](http://dev.office.com/reference/add-ins/shared/settings)|コンテンツおよび作業ウィンドウ|アドインが連携しているドキュメント、スプレッドシート、またはプレゼンテーション。コンテンツおよび作業ウィンドウのアドイン設定は、その設定が保存されているドキュメントから、その設定を作成したアドインで使用できます。**重要:****Settings** オブジェクトを使用して、パスワードおよびその他の機密の個人を特定できる情報 (PII) を保存しないでください。保存されたデータはエンド ユーザーに対して表示されませんが、ドキュメントの一部として保存されているため、ドキュメントのファイル形式を直接読み取ることでアクセスできます。アドインによる PII の使用と、アドインが必要とするすべての PII の保存は、開発するアドインをユーザーのセキュリティが保護されるリソースとしてホストするサーバーのみで行うよう制限する必要があります。|Word、Excel、または PowerPoint **メモ:** Project 2013 の作業ウィンドウ アドインでは、アドインの状態または設定を保存するための **Settings** API をサポートしていません。ただし、Project (および他の Office ホスト アプリケーション) で動作するアドインの場合は、ブラウザーの Cookie や Web ストレージなどのテクニックを使用することができます。このようなテクニックの詳細については、「[Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)」を参照してください。 |
|[RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings)|Outlook|アドインがインストールされている、ユーザーの Exchange サーバー メールボックス。これらの設定はユーザーのサーバー メールボックスに保存されるので、ユーザーと共に "ローミング" でき、そのユーザーのメールボックスにアクセスしている、サポートされているクライアント ホスト アプリケーションまたはブラウザーのコンテキストでアドインが実行されている場合、そのアドインでこれらの設定を利用できます。 Outlook アドインのローミング設定は、その設定を作成したアドインのみが利用でき、また、アドインがインストールされているメールボックスからのみ利用できます。|Outlook|
|[CustomProperties](http://dev.office.com/reference/add-ins/outlook/CustomProperties)|Outlook|アドインが連携するメッセージ、予定、または会議出席依頼アイテム。 Outlook アドイン アイテムのカスタム プロパティは、そのプロパティを作成したアドインのみが利用でき、また、プロパティが保存されているアイテムからのみ利用できます。|Outlook|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a>実行時のメモリ内での設定データの管理


内部的には、 **Settings** オブジェクト、 **CustomProperties** オブジェクト、または **RoamingSettings** オブジェクトでアクセスされるプロパティ バッグ内のデータは、名前/値のペアを含むシリアル化された JavaScript Object Notation (JSON) オブジェクトとして格納されます。各値の名前 (キー) は **string** である必要があり、格納された値は JavaScript の **string**、 **number**,  **date**、または  **object** にすることが可能ですが、 **function** にすることはできません。

この例はプロパティ バッグの構造を示し、3 つの定義された  **string** 値 ( `firstName`、 `location`、および  `defaultView` という名前) が含まれます。




```
{
"firstName":"Erik",
"location":"98052",
"defaultView":"basic"
}
```

前のアドイン セッションで設定プロパティ バッグが保存されると、アドインが初期化されるとき、またはその後のアドインの現在のセッション中の任意の時点で、その設定プロパティ バッグを読み込むことができます。セッションの間、設定は、作成している設定の種類に対応するオブジェクト ( **Settings**、 **CustomProperties**、または  **RoamingSettings**) の  **get**、 **set**、および  **remove** メソッドを使用して、全体がメモリ内で管理されます。 


 >**重要**  アドインの現在のセッションの間に行われた追加、更新、または削除を保存場所に保持するには、その種の設定を操作する際に使用する、対応するオブジェクトの  **saveAsync** メソッドを呼び出す必要があります。 **get**、 **set**、および  **remove** メソッドは、設定プロパティ バッグのメモリ内コピーにのみ作用します。 **saveAsync** の呼び出しなしにアドインが閉じられた場合、そのセッションの間に設定に対して行われた変更は失われます。 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a>コンテンツ アドインおよび作業ウィンドウ アドインで、ドキュメントごとにアドインの状態と設定を保存する方法


Word、Excel、または PowerPoint 用のコンテンツ アドインまたは作業ウィンドウ アドインの状態またはカスタム設定を保持するには、[Settings](http://dev.office.com/reference/add-ins/shared/settings) オブジェクトとそのメソッドを使用します。**Settings** オブジェクトのメソッドを使用して作成されたプロパティ バッグは、それを作成したコンテンツ アドインまたは作業ウィンドウ アドインのインスタンスのみが利用でき、プロパティ バッグが保存されているドキュメント以外からは使用できません。

**Settings** オブジェクトは、[Document](http://dev.office.com/reference/add-ins/shared/document) オブジェクトの一部として自動的に読み込まれ、作業ウィンドウ アドインまたはコンテンツ アドインがアクティブ化されると使用できるようになります。**Document** オブジェクトがインスタンス化された後は、**Document** オブジェクトの [settings](../../reference/shared/document.settings.md) プロパティを使用して、**Settings** オブジェクトにアクセスできます。セッションの存続中は、**Settings.get**、**Settings.set**、および **Settings.remove** メソッドを使用するだけで、永続的な設定およびアドインの状態の読み取り、書き込み、または削除をプロパティ バッグのメモリ内コピーで行うことができます。

set メソッドと remove メソッドは設定プロパティ バッグのメモリ内コピーに対してのみ動作するので、アドインが関連付けられているドキュメントに新しい設定を保存、または変更された設定を保存し直すには [Settings.saveAsync](../../reference/shared/settings.saveasync.md) メソッドを呼び出す必要があります。


### <a name="creating-or-updating-a-setting-value"></a>設定値の作成または更新

次のコード例では、[Settings.set](../../reference/shared/settings.set.md) メソッドを使用して `'themeColor'` という名前の設定を作成し、値 `'green'` を指定する方法を説明します。set メソッドの最初のパラメーターは、設定するか作成する設定の _name_ (Id) であり、これは大文字と小文字が区別されます。2 番目のパラメーターは、設定の _value_ です。


```
Office.context.document.settings.set('themeColor', 'green');
```

 指定した名前を持つ設定は、それがまだ存在していない場合には作成され、すでに存在している場合はその値が更新されます。**Settings.saveAsync** メソッドを使用すると、新しい設定または更新された設定をドキュメントに保持できます。


### <a name="getting-the-value-of-a-setting"></a>設定値の取得

次の例では、[Settings.get](../../reference/shared/settings.get.md) メソッドを使用して "themeColor" という名前の設定値を取得する方法を示します。**get** メソッドの唯一のパラメーターは、設定の _name_ であり、これは大文字と小文字が区別されます。


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 **get** メソッドでは、指定した _name_ という設定に対して以前に保存した値を返します。設定が存在しない場合、メソッドは **null** を返します。


### <a name="removing-a-setting"></a>設定の削除

次の例では、[Settings.remove](../../reference/shared/settings.removehandlerasync.md) メソッドを使用して、"themeColor" という名前の設定を削除する方法を示します。**remove** メソッドの唯一のパラメーターは設定の _name_ であり、これは大文字と小文字が区別されます。


```
Office.context.document.settings.remove('themeColor');
```

該当する設定が存在しない場合は何も起きません。ドキュメントから設定を削除したままにする場合は、 **Settings.saveAsync** メソッドを使用します。


### <a name="saving-your-settings"></a>設定の保存

現在のセッション中に、アドインがメモリ内の設定プロパティ バッグに対して行った追加、変更、または削除を保存するには、[Settings.saveAsync](../../reference/shared/settings.saveasync.md) メソッドを呼び出してそれらの設定をドキュメントに保存する必要があります。**saveAsync** メソッドの唯一のパラメーターは _callback_ であり、これはパラメーターを 1 つだけ取るコールバック関数です。 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

**saveAsync** メソッドに _callback_ パラメーターとして渡した匿名関数は、操作の完了時に実行されます。コールバックの _asyncResult_ パラメーターは、処理のステータスを含む **AsyncResult** オブジェクトへのアクセスを提供します。例では、関数は **AsyncResult.status** プロパティを調べて、保存操作が成功したのか失敗したのかを確認し、アドインのページにその結果を表示します。


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Outlook アドインでユーザーのメールボックスに設定をローミング設定として保存する方法


Outlook アドインは、 [RoamingSettings](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) オブジェクトを使用して、ユーザーのメールボックスに固有の、アドインの状態および設定のデータを保存できます。このデータには、アドインを実行しているユーザーではなく、Outlook アドインのみがアクセスできます。データはユーザーの Exchange Server メールボックスに格納されます。データには、ユーザーが自分のアカウントにログインして Outlook アドインを実行したときにアクセスできるようになります。


### <a name="loading-roaming-settings"></a>ローミング設定の読み込み


通常、Outlook アドインでは、 [Office.initialize](../../reference/shared/office.initialize.md) イベント ハンドラーでローミング設定を読み込みます。次の JavaScript のコード例は、既存のローミング設定を読み込む方法を示しています。


```
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a>ローミング設定の作成または割り当て


前の例に続けて、次の  `setAppSetting` 関数では、 [RoamingSettings.set](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) メソッドを使用して、 `cookie` という名前の設定項目に今日の日付を設定、または今日の日付で更新する方法を示しています。次に、 [RoamingSettings.saveAsync](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) メソッドを使用して Exchange Server にすべてのローミング設定を保存し直しています。


```
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

**saveAsync** メソッドは、ローミング設定を非同期で保存し、オプションのコールバック関数を受け取ります。このコード例では、`saveMyAppSettingsCallback` という名前のコールバック関数を **saveAsync** メソッドに渡します。非同期呼び出しが返ると、`saveMyAppSettingsCallback` 関数の _asyncResult_ パラメーターが [AsyncResult](http://dev.office.com/reference/add-ins/outlook/simple-types) オブジェクトにアクセスします。このオブジェクトを使用すると、**AsyncResult.status** プロパティで操作の成功または失敗を判定することができます。


### <a name="removing-a-roaming-setting"></a>ローミング設定の削除


また、次の  `removeAppSetting` 関数は、前の例をさらに拡張するものです。この例では、 [RoamingSettings.remove](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) メソッドを使用して `cookie` 設定を削除し、すべてのローミング設定を Exchange Server に保存し直す方法を示しています。


```
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Outlook アドインでアイテムごとに設定をカスタムプロパティとして保存する方法


カスタム プロパティを使用すると、Outlook アドインは処理しているアイテムに関する情報を保存できます。たとえば、Outlook アドインを使用して、メッセージ内の会議の提案から予定を作成する場合は、カスタム プロパティを使用して、会議が作成されたという事実を保存できます。これにより、メッセージを再び開いたときに、Outlook アドインが再び予定の作成を行うことはありません。

メッセージ、予定、または会議出席依頼の特定のアイテムに対してカスタム プロパティを使用するには、その前に、 [Item](../../reference/outlook/Office.context.mailbox.item.md) オブジェクトの **loadCustomPropertiesAsync** メソッドを呼び出して、プロパティをメモリに読み込む必要があります。現在のアイテムに対してカスタム プロパティが既に設定されている場合は、この時点で Exchange サーバーから読み込まれます。プロパティを読み込んだ後、 [CustomProperties](http://dev.office.com/reference/add-ins/outlook/CustomProperties) オブジェクトの [set](http://dev.office.com/reference/add-ins/outlook/RoamingSettings) メソッドおよび **get** メソッドを使用して、メモリ内のプロパティの追加、更新、および取得を実行できます。アイテムのカスタム プロパティに対して行った変更を保存するには、 [saveAsync](http://dev.office.com/reference/add-ins/outlook/CustomProperties) メソッドを使用して、アイテムに加えた変更を Exchange サーバー上で保持する必要があります。


### <a name="custom-properties-example"></a>カスタム プロパティの例

以下の例では、カスタム プロパティを使用する Outlook アドインの一連の関数を、簡略化して示しています。この例を出発点として、カスタム プロパティを使用する Outlook アドインを作成できます。 

これらの関数を使用する Outlook アドインは、次の例に示すように、 `_customProps` 変数で **get** メソッドを呼び出すことによって、任意のカスタム プロパティを取得します。




```
var property = _customProps.get("propertyName");
```

以下の例には、次の関数が含まれています。



|**関数名**|**説明**|
|:-----|:-----|
| `Office.initialize`|アドインを初期化し、Exchange サーバーから現在のアイテムのカスタム プロパティを読み込みます。|
| `customPropsCallback`|Exchange サーバーから返されるカスタム プロパティを取得し、後で使用できるように保存します。|
| `updateProperty`|特定のプロパティを設定または更新し、その変更を Exchange サーバーに保存します。|
| `removeProperty`|特定のプロパティを削除し、その削除を Exchange サーバーに保存します。|
| `saveCallback`|`updateProperty` 関数および `removeProperty` 関数内で **saveAsync** メソッドを呼び出すためのコールバック|



```
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change 
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal 
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method. 
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```


## <a name="additional-resources"></a>その他のリソース



- [JavaScript API for Office について](../../docs/develop/understanding-the-javascript-api-for-office.md)
    
- [Outlook アドイン](../outlook/outlook-add-ins.md)
    
- [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
