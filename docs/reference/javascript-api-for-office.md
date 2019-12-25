---
layout: LandingPage
ms.topic: landing-page
title: Office JavaScript API リファレンス ドキュメント
description: Office JavaScript Api について説明します。
ms.date: 12/24/2019
localization_priority: Priority
ms.openlocfilehash: c10eeb5c89a74b28e9af44bf72b20a7ad610738b
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: ja-JP
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851552"
---
# <a name="api-reference-documentation"></a>API リファレンス ドキュメント

アドインは Office JavaScript API を使用することで、Office ホスト アプリケーション内のオブジェクトを操作できます。 

<ul>
    <li><b>ホスト固有</b> API では、特定の Office アプリケーションにネイティブなオブジェクトを操作するために使用できる、厳密に型指定されたオブジェクトが提供されます。</li>
    <li><b>共通 API</b> を使用すると、複数の種類の Office アプリケーション間で共通の UI、ダイアログ、クライアント設定などの機能にアクセスすることができます。</li>
</ul>

可能な場合は常にホスト固有 API を使用するようにし、ホスト固有 API でサポートされていないシナリオに対してのみ共通 API を使用するようにします。 これらの 2 つの API モデルの詳細については、「<a href="../overview/office-add-ins-fundamentals.md#api-models">Office アドインの構築</a>」を参照してください。

<h2>API リファレンス</h2>

<ul class="panelContent cardsF cols cols3">
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/excel"><img src="../images/index/logo-excel.svg" alt="Excel API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Excel API リファレンス</h3>
                        <p><a href="/javascript/api/excel">Excel アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/outlook"><img src="../images/index/logo-outlook.svg" alt="Outlook API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Outlook API リファレンス</h3>
                        <p><a href="/javascript/api/outlook">Outlook アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/word"><img src="../images/index/logo-word.svg" alt="Word API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>Word API リファレンス</h3>
                        <p><a href="/javascript/api/word">Word アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/powerpoint"><img src="../images/index/logo-powerpoint.svg" alt="PowerPoint API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>PowerPoint API リファレンス</h3>
                        <p><a href="/javascript/api/powerpoint">PowerPoint アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/onenote"><img src="../images/index/logo-onenote.svg" alt="OneNote API reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>OneNote API リファレンス</h3>
                        <p><a href="/javascript/api/onenote">OneNote アドイン構築用の JavaScript APIs。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
    <li>
        <div class="cardSize">
            <div class="cardPadding">
                <div class="card">
                    <div class="cardImageOuter">
                        <div class="cardImage">
                            <a href="/javascript/api/office"><img src="../images/index-landing-page/i_code-blocks.svg" alt="reference docs" /></a>
                        </div>
                    </div>
                    <div class="cardText">
                        <h3>共通 API リファレンス</h3>
                        <p><a href="/javascript/api/office">すべての Office アドインで使用できる JavaScript API。</a></p>
                    </div>
                </div>
            </div>
        </div>
    </li>
</ul>

<b>注</b>: 現在、Project 用のホスト固有 API はありません。Project アドインを構築する場合は、共通 API を使用してください。また、PowerPoint 用のホスト固有 API の範囲は非常に限定的であるため、PowerPoint アドインを構築する際は、主に共通 API を使用してください。

<h2>Open API の仕様</h2>

新しい Office アドイン用の API の設計と開発にあたり、[Open API の仕様](openspec/openspec.md) ページでこれらに対するフィードバックの提供が可能になります。パイプラインの新機能をご確認いただき、設計の仕様に関する情報をお寄せください。