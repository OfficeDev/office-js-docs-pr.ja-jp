
# <a name="diagnostics"></a>診断

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics

Outlook アドインに診断情報を提供します。

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

### <a name="members"></a>メンバー

####  <a name="hostname-string"></a>ホスト名: 文字列

ホスト アプリケーションの名前を表す文字列を取得します。

文字列は、`Outlook`、`OutlookIOS`、または `OutlookWebApp` のいずれかの値になります。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

####  <a name="hostversion-string"></a>ホストバージョン: 文字列

ホスト アプリケーションまたは Exchange Server のバージョンを表す文字列を取得します。

メール アドインを Outlook デスクトップ クライアントまたは Outlook for iOS で実行している場合、`hostVersion` プロパティは、ホスト アプリケーションである Outlook のバージョンを返します。Outlook Web App では、プロパティは、Exchange Server のバージョンを返します。たとえば、文字列 `15.0.468.0` です。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|

####  <a name="owaview-string"></a>OWAView: 文字列

Outlook Web App の現在のビューを表す文字列を取得します。

返される文字列は、`OneColumn` 、`TwoColumns`、または `ThreeColumns` のいずれかの値になります。

ホスト アプリケーションが Outlook Web App ではない場合、このプロパティへのアクセスは `undefined` となります。

Outlook Web App には、画面の幅、ウィンドウの幅、および表示可能な列数に応じて 3 つのビューがあります。

*   `OneColumn`これは、画面の幅が狭い場合に表示されます。Outlook Web App は、このシングル コラム レイアウトを使用してスマートフォンの画面全体に表示します。
*   `TwoColumns`これは、画面の幅がより広い場合に表示されます。Outlook Web App は、ほとんどのタブレットでこのビューを使用します。
*   `ThreeColumns`これは、画面の幅が広い場合に表示されます。Outlook Web App は、デスクトップ コンピュータの全画面表示ウィンドウなどでこのビューを使用します。

##### <a name="type"></a>型:

*   文字列

##### <a name="requirements"></a>要件

|要件| 値|
|---|---|
|[メールボックスの最低要件セットのバージョン](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[最小限のアクセス許可レベル](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[適用可能な Outlook のモード](https://docs.microsoft.com/outlook/add-ins/#extension-points)| 新規作成または読み取り|