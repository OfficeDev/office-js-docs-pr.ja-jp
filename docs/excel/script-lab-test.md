---
title: Script Lab 統合のテスト
description: ''
ms.date: 12/04/2017
---


# <a name="testing-script-lab-integration"></a>Script Lab 統合のテスト

間もなく使用可能になる ScriptLab 機能をデモンストレーションするためのサンプル テスト ファイルです。ScriptLab 機能を使用すると、開発者は Excel、Word、PowerPoint でスニペットを試すことができます。  

## <a name="prerequisites"></a>前提条件
- ScriptLab スニペットのビュー URL が必要になります。
- 注: ScriptLab では、最新のスニペットを探索するために Office 365 が必要であることが*示されるはず*です。開発者は、開発目的に限定して [Office 365 開発者プログラム](https://dev.office.com/devprogram)から Office 365 サブスクリプションを取得できます。  


## <a name="try-it-out-button"></a>試してみる 'ボタン'
この方法で、「試してみるボタン」を追加します。このボタンにより、推奨事項がコード スニペットと関連付けられます。これを可能にするため、Office UI Fabric クラスを使用して、リンクをボタンとしてスタイル設定します。リンク自体で、*aria label* 属性を設定してください。

**デモ:**

<a href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">試してみる</a>


<button href="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">試してみる</button>


**コード:**
```html
<a href="ahttps://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" class="ms-Button" aria-label="Open this snippet in Script Lab, an Office Add-in">Try it out</a>
```



## <a name="embed-script-lab-as-an-iframe"></a>Script Lab を iframe として埋め込む
このモードでは、スニペットを、iframe として直接ドキュメントに埋め込みます。幅は (他のすべてのスニペットの幅に基づいて) 95% に設定されています。iframe のフレーム境界線を削除することをお勧めします。通常、高さはスニペットに合わせて調整する必要があります。

**デモ:**
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>

**コード:**
```html
<iframe src="https://bornholm-edge.azurewebsites.net/#/view/gist/excel/0cc24cee687141d1c2726c0feea70911" height="600px" width="95%" frameborder="0"></iframe>
```

## <a name="testing-considerations"></a>テストに関する考慮事項
モバイルの Office 365 以外のサブスクリプションを検証する必要があります (多くの開発者が 2013 以前の office js ドキュメントを使用しているとのフィードバックがあります)。  

埋め込みパスに関して、最終的な承認が必要になります。また、view gist ページに公開されるコンテンツがユーザー補助ガイドラインを満たしていることを確認することも必要です。

## <a name="see-also"></a>関連項目
