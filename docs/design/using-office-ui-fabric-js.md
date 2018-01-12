
# <a name="use-office-ui-fabric-js-in-office-add-ins"></a>Office アドインでの Office UI Fabric JS の使用

Office UI Fabric は、Office と Office 365 のユーザー エクスペリエンスを構築するための JavaScript フロント エンドのフレームワークです。Angular や React などのフレームワークを使用せず、JavaScript のみを使用してアドインをビルドする場合は、ユーザー エクスペリエンスを作成するために Fabric JS の使用を検討してください。詳細については、「[Office UI Fabric JS](https://dev.office.com/fabric-js)」を参照してください。

この記事では、Fabric JS の基本的な使用方法について説明します。  

## <a name="add-the-fabric-cdn-references"></a>Fabric CDN 参照の追加
CDN から Fabric を参照するには、次に示す HTML コードをページに追加します。

    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.min.css">
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/css/fabric.components.min.css">
    <script src="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-js/1.4.0/js/fabric.min.js"></script>

## <a name="use-fabric-js-ux-components"></a>Fabric JS UX コンポーネントの使用

Fabric JS は、アドインで使用できるボタンやチェックボックスなど、複数の UX コンポーネントを提供しています。次に、アドインでの使用をお勧めする Fabric JS UX コンポーネントのリストを示します。アドインで Fabric コンポーネントのいずれかを使用するには、その Fabric のドキュメントへのリンクをたどって、「**このコンポーネントの使用方法**」の手順を実行してください。 

- [Breadcrumb](https://dev.office.com/fabric-js/Components/Breadcrumb/Breadcrumb.html)
- [Button](https://dev.office.com/fabric-js/Components/Button/Button.html) (アドインで小さなボタンのバリエーションの使用を検討してください。タッチ デバイスで最小 40px のタッチ ターゲットを確保するために、小さいボタンに 16px のパディングを追加します。)
- [Checkbox](https://dev.office.com/fabric-js/Components/CheckBox/CheckBox.html)
- [ChoiceFieldGroup](https://dev.office.com/fabric-js/Components/ChoiceFieldGroup/ChoiceFieldGroup.html)
- [Date Picker](https://dev.office.com/fabric-js/Components/DatePicker/DatePicker.html) (アドインに日付ピッカーを実装する方法の例は、[Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) コード サンプルを参照してください)
- [Dropdown](https://dev.office.com/fabric-js/Components/Dropdown/Dropdown.html)
- [Label](https://dev.office.com/fabric-js/Components/Label/Label.html)
- [Link](https://dev.office.com/fabric-js/Components/Link/Link.html)
- [List](https://dev.office.com/fabric-js/Components/List/List.html) (コンポーネントの既定のスタイルを CSS で変更することを検討してください)
- [MessageBanner](https://dev.office.com/fabric-js/Components/MessageBanner/MessageBanner.html)
- [MessageBar](https://dev.office.com/fabric-js/Components/MessageBar/MessageBar.html)
- [Overlay](https://dev.office.com/fabric-js/Components/Overlay/Overlay.html)
- [Panel](https://dev.office.com/fabric-js/Components/Panel/Panel.html)
- [Pivot](https://dev.office.com/fabric-js/Components/Pivot/Pivot.html)
- [ProgressIndicator](https://dev.office.com/fabric-js/Components/ProgressIndicator/ProgressIndicator.html)
- [Searchbox](https://dev.office.com/fabric-js/Components/SearchBox/SearchBox.html)
- [Spinner](https://dev.office.com/fabric-js/Components/Spinner/Spinner.html)
- [Table](https://dev.office.com/fabric-js/Components/Table/Table.html)
- [TextField](https://dev.office.com/fabric-js/Components/TextField/TextField.html)
- [Toggle](https://dev.office.com/fabric-js/Components/Toggle/Toggle.html)
   
## <a name="updating-your-add-in-to-use-fabric-js"></a>Fabric JS を使用するためのアドインの更新
以前のバージョンの Office UI Fabric を使用しており、Fabric JS への移行を考えている場合は、新しいコンポーネントをアドインに組み込んでテストする方法について理解していることが必要です。更新を計画するには、次に示す点に留意してください。

- Fabric JS を使用することでコンポーネントの初期化が簡単になります。以前のバージョンの Fabric の場合は、Fabric コンポーネントの JavaScript ファイルをアドイン プロジェクト (そのファイルへの `<Script>` 参照が含まれているプロジェクト) に含めてからコンポーネントを初期化します。Fabric JS では、Fabric コンポーネントの JavaScript ファイルと、それに関連する `<Script>` 参照を含める必要はなくなりました。Fabric コンポーネントの初期化以外に必要な手順はありません。   
- いくつかのコンポーネントは、UX コンポーネントの動作を制御する関数を提供するようになりました。たとえば、チェックボックス コントロールには、チェックボックスのオン状態とオフ状態を切り替える `toggle` 関数があります。 
- 一部のアイコン クラスの名前とスタイルが更新されています。
- 最重要の変更点は、多数のコンポーネントで `<label>` 要素を使用していることです。`<label>` 要素では、コンポーネントのスタイルを制御します。`<label>` 要素を使用するように UX コードを更新することが必要になる場合があります。たとえば、Fabric JS のチェック ボックスがオンになっている `<input>` 要素の属性の値を変更しても、そのチェック ボックスに影響はありません。代わりに、`check`、`unCheck`、`toggle` の関数をお使いください。   

## <a name="next-steps"></a>次の手順
Fabric JS の使用方法を示す完全なコード サンプルをご用意しています。次に示すリソースをご覧ください。

- [Excel Sales Tracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) 

## <a name="related-resources"></a>関連リソース
以前のリリースの Fabric に関するコード サンプルやドキュメントを探している場合は、次に示す記事をご確認ください。

- [UX 設計パターン (Fabric 2.6.1 を使用)](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code) 
- [Office アドイン Fabric UI サンプル (Fabric 1.0 を使用)](https://github.com/OfficeDev/Office-Add-in-Fabric-UI-Sample) 
- [Office アドインでの Fabric 2.6.1 の使用](https://dev.office.com/docs/add-ins/design/ui-elements/using-office-ui-fabric)
 

