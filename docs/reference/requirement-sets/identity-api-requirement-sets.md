# <a name="identity-api-requirement-sets"></a>Identity API の要件セット

要件セットは、API メンバーの名前付きグループです。 Office アドインでは、マニフェストで指定されている要件のセットを使用して、またはランタイム チェックを使用して、Office ホストがアドインを必要とする Api をサポートしているかどうかを決定します。 詳細については、 [Office のバージョンおよび要件の設定](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)を参照してください。

Office アドインは、Office の複数のバージョン間で実行します。 アイデンティティの API の要件のセット、Office アプリケーションの要件のセット、およびビルドまたはバージョン番号をサポートする Office ホスト アプリケーションを次の表に一覧します。

|  要件セット  | Office 2013 for Windows | Office for Windows   |  IPad の office 365  |  Office for Mac  | Office Online  | SharePoint Online | OneDrive.com |Outlook.com および Exchange Online|
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| IdentityAPI 1.1  | 該当なし | プレビュー ***** | 近日公開 | プレビュー *****| プレビュー | プレビュー| 近日公開 | 近日公開 |

> ***** プレビュー段階では、識別情報 API は、2016 の Windows と Mac の高速オプションを使用して内部関係者がプログラムのユーザーに対してのみサポートされています。 内部関係者がプログラムに参加するには、 [Office 内部者をする](https://products.office.com/office-insider?tab=tab-1)を参照してください。 ファスト ・ トラックを切り替えるには、 [内部関係者による高速](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_officeinsider-mso_win10-msoinsider_reg/its-here-office-insider-fast-for-office-2016-on/dbe8e7bb-9523-44a4-948b-9436fedfd961)参照してください。

バージョン、ビルド番号、Office Online Server の詳細については以下を参照してください。

- [Office 365 クライアントの更新プログラム チャネル リリースのバージョン番号およびビルド番号](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [使用している Office のバージョンを確認する方法](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Office 365 クライアント アプリケーションのバージョン番号およびビルド番号を確認することができます。](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Office Online Server 概要](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Office 共通 API の要件セット

共通 API の要件セットの詳細は、「[Office 共通 API の要件セット](office-add-in-requirement-sets.md)」を参照してください。

## <a name="identityapi-11"></a>IdentityAPI 1.1 

シングル サインオン IdentityAPI 1.1 は API の最初のバージョンです。 この API についての詳細は、 [アドインで SSO を有効にする](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins) の [SSO API リファレンス](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)を参照してください。

## <a name="see-also"></a>関連項目

- [Office のバージョンと要件セット](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Office のホストと API の要件を指定する](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [Office アドインの XML マニフェスト](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
