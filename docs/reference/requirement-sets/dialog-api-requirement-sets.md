# <a name="dialog-api-requirement-sets"></a>ダイアログ API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件セットを使用して、またはランタイム チェックを使用して、Office ホストがアドインを必要とする API をサポートするかどうかを決定します。 詳細については、 「 [Office のバージョンおよび要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)」を参照してください。

Office アドインは Office の複数のバージョンで機能します。次の表は、ダイアログ API の要件セット、その要件セットをサポートする Office ホスト アプリケーション、Office アプリケーションのビルド番号またはバージョン番号の一覧です。

|  要件セット  | Office 2013 for Windows | Office 2016 for Windows (MSI インストール )   | Office 365 for Windows (C2R インストール )   |  Office 365 for iPad  |  Office for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| DialogApi 1.1  | ビルド 15.0.4855.1000 以降 | ビルド 16.0.4390.1000 以降 | バージョン 1602 ( ビルド 6741.0000) 以降 | 1.22 以降 | 15.20 以降| 2017 年 1 月 | バージョン 1608 ( ビルド 7601.6800) 以降|

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [次で Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細は、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」を参照してください。

## <a name="dialog-api-11"></a>ダイアログ API 1.1 

ダイアログ API の 1.1 は、API の最初のバージョンです。 API についての詳細は、「[ダイアログ  API ](/javascript/api/office/office.ui)」の参照トピックを参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office ホストと API 要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
