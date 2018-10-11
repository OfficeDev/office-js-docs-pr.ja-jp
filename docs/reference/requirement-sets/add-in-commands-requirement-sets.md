# <a name="add-in-commands-requirement-sets"></a>アドイン コマンドの要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件セットを使用して、またはランタイム チェックを使用して、Office ホストがアドインを必要とする API をサポートするかどうかを決定します。 詳細については、 「 [Office のバージョンおよび要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。

アドイン コマンドは、Office UI を拡張し、アドインでアクションを開始する UI 要素です。アドイン コマンドを使用すると、リボン上のボタンやアイテムをコンテキスト メニューに追加できます。詳細については、「[Excel、Word および PowerPoint のアドイン コマンド](https://docs.microsoft.com/office/dev/add-ins/design/add-in-commands)」と「[Outlook のアドイン コマンド](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)」を参照してください。

アドイン コマンドの最初のリリースには、対応する要件セットがありません (つまり、AddinCommands 1.0 要件セットはありません)。次の表に、初期リリースのバージョンをサポートする Office ホスト アプリケーション、およびそれらのアプリケーションのビルド バージョンまたはビルド番号を示します。  

| リリース   |  Office 2013 for Windows | Office 2016 for Windows (非サブスクリプション) | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| アドイン コマンド (初期リリース、要件設定なし) | 該当なし | 16.0.4678.1000 *Outlook のみでサポートされています。* |バージョン 1603 (ビルド 6769.0000) 以降 | 該当なし | 15.33 以降| 2016 年 1 月 | |

アドイン コマンド 1.1 の要件セットでは、「[ドキュメントで作業ウィンドウを自動的に開く](https://docs.microsoft.com/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document)」機能が導入されています。

次の表に、アドイン コマンド 1.1 の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号を示します。 

|  要件セット  |  Office 2013 for Windows | Office 2016 for Windows (非サブスクリプション) | Office 365 for Windows   |  Office 365 for iPad  |  Office 365 for Mac  | Office Online  |  
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| AddInCommands 1.1  | 該当なし | 16.0.4678.1000 *Outlook のみでサポートされています。*  | バージョン 1705 (ビルド 8121.1000) 以降 | 該当なし | 15.34 以降| 2017 年 5 月 | |

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [次で Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細は、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」を参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office ホストと API 要件を指定](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
