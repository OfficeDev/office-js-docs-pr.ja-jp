# <a name="define-add-in-commands-in-your-manifest"></a>Outlook 用のマニフェストでアドイン コマンドを定義する

アドイン コマンドは、操作を実行する UI 要素を使用して、既定の Office UI をカスタマイズする簡単な方法を提供します。たとえば、リボンにカスタムのボタンを追加できます。コマンドを作成する場合は、既存のマニフェストに **[VersionOverrides](../../reference/manifest/versionoverrides.md)** ノードを追加します。 

マニフェストが **VersionOverrides** 要素を含む場合、アドイン コマンドをサポートする Word、Excel、Outlook、PowerPoint のバージョンは、要素内の情報を使用して、アドインをロードします。アドイン コマンドをサポートしていない以前のバージョンの Office 製品では、要素は無視されます。

クライアント アプリケーションが **VersionOverrides** ノードを認識する場合、アドインの名前はリボンに表示され、読み込み/作成ウィンドウには表示されません。これらの場所に、アドインは表示されません。
 
## <a name="versionoverrides"></a>VersionOverrides

[VersionOverrides](../../reference/manifest/versionoverrides.md) 要素は、アドインによって実装されたアドイン コマンドに関する情報を格納するルート要素です。これは、マニフェスト スキーマ v1.1 以降でサポートされています。

**VersionOverrides** スキーマには 2 つのバージョンがあります。

| スキーマのバージョン | 説明 |
|----------------|-------------|
| 1.0 | Office アプリのデスクトップ バージョンのアドイン コマンドをサポートします。 | 
| 1.1 | [ピン留め可能な作業ウィンドウ](https://docs.microsoft.com/outlook/add-ins/pinnable-taskpane)およびモバイル アドインのサポートを追加します。**注:**現時点では、Outlook 2016 for Windows および Outlook for iOS でのみサポートされています。 |

アドインは、より新しいバージョンを前のバージョンの内側に入れ子にすることにより、**VersionOverrides** スキーマの複数のバージョンをサポートできます。これにより、クライアントがより新しいバージョンをサポートして新機能を利用できるようにしつつ、前のクライアントが旧バージョンを読み込めるようにします。詳細については、「[複数のバージョンを実装する](../../reference/manifest/versionoverrides.md#implementing-multiple-versions)」を参照してください。

**VersionOverrides** 要素には、次の子要素が含まれます。

- [Description](../../reference/manifest/description.md)
- [Requirements](../../reference/manifest/requirements.md)
- [Hosts](../../reference/manifest/hosts.md)
- [Resources](../../reference/manifest/resources.md)
- [VersionOverrides](../../reference/manifest/versionoverrides.md)

次の図は、アドイン コマンドの定義に使用する要素の階層を示しています。 

![マニフェスト内のアドイン コマンド要素の階層](../../images/080da303-51c4-4882-b74a-7ba11517c0ad.png)

## <a name="sample-manifests"></a>マニフェストのサンプル

Word、Excel、PowerPoint でアドイン コマンドを実装するサンプル マニフェストの場合は、「[Simple add-in commands sample](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/tree/master/Simple)」 (簡単なアドイン コマンドのサンプル) をご覧ください。

Outlook でアドイン コマンドを実装するサンプル マニフェストの場合は、「[Sample manifest file for an Outlook add-in](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml)」 (Outlook アドイン用のサンプル マニフェスト ファイル) をご覧ください。

## <a name="additional-resources"></a>その他のリソース

- [Outlook のアドイン コマンド](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)
    
- [Outlook アドインのマニフェスト](https://docs.microsoft.com/outlook/add-ins/manifests)
    
- [Outlook アドイン コマンドのデモ サンプル](https://github.com/OfficeDev/outlook-add-in-command-demo)
