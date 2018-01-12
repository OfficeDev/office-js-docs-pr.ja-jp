# <a name="install-the-latest-version-of-office-2016"></a>Office 2016 の最新バージョンをインストールする

新しい開発者用機能 (現時点ではプレビュー版のものを含む) は、Office の最新ビルドの取得をオプトインしたサブスクライバーに最初に配信されます。Office 2016 の最新ビルドをオプトインするには: 

- Office 365 Home、Personal、または University のサブスクライバーは、「[Office Insider プログラム](https://products.office.com/en-us/office-insider)」を参照してください。
- 一般法人向け Office 365 をご利用の場合は、「[一般法人向け Office 365 の先行リリース ビルドをインストールする](https://support.office.com/en-us/article/Install-the-First-Release-build-for-Office-365-for-business-customers-4dd8ba40-73c0-4468-b778-c7b744d03ead?ui=en-US&rs=en-US&ad=US)」を参照してください。
- Mac で Office 2016 を実行している場合:
    - Office 2016 for Mac プログラムを起動します。
    - [ヘルプ] メニューで [**更新プログラムのチェック**] を選びます。
    - [Microsoft AutoUpdate] ボックスで、チェック ボックスをオンにして Office Insider プログラムに参加します。 

最新ビルドを取得するには: 

1. [Office 2016 展開ツール](https://www.microsoft.com/en-us/download/details.aspx?id=49117)をダウンロードします。 
2. ツールを実行します。これにより、Setup.exe および configuration.xml という 2 つのファイルが抽出されます。
3. configuration.xml を[先行リリース構成ファイル](https://raw.githubusercontent.com/OfficeDev/Office-Add-in-Commands-Samples/master/Tools/FirstReleaseConfig/configuration.xml)に置き換えます。
4. 次のコマンドを管理者として実行します: `setup.exe /configure configuration.xml` 

>**注:**このコマンドの実行には時間がかかることがあります (進行状況は表示されません)。

インストール処理の完了時点で、最新の Office 2016 アプリケーションがインストールされています。最新のビルドであることを確認するには、任意の Office アプリケーションから **[ファイル]**  >  **[アカウント]** に移動します。[Office 更新プログラム] に、[(Office Insiders)] ラベルが表示され、その下にバージョン番号が表示されます。

![Office Insiders のラベルと製品情報を示すスクリーンショット](../../images/officeinsider.PNG)

## <a name="minimum-office-builds-for-office-javascript-api-requirement-sets"></a>Office JavaScript API の要件セットの最小 Office ビルド

API の要件セットの各プラットフォームの最小製品ビルドについては、次をご覧ください。

- [Word JavaScript API の要件セット](../../reference/requirement-sets/word-api-requirement-sets.md)
- [Excel JavaScript API の要件セット](../../reference/requirement-sets/excel-api-requirement-sets.md)
- [OneNote JavaScript API の要件セット](../../reference/requirement-sets/onenote-api-requirement-sets.md)
- [ダイアログ API の要件セット](../../reference/requirement-sets/dialog-api-requirement-sets.md)
- [Office 共通 API の要件セット](../../reference/requirement-sets/office-add-in-requirement-sets.md)
