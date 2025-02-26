document.addEventListener("DOMContentLoaded", function () {
    document.getElementById("simplifyButton").addEventListener("click", simplifyEmail);
});

async function simplifyEmail() {
    await Office.onReady();
    Office.context.mailbox.item.body.getAsync("text", function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            let emailText = result.value;
            let useBenefitLanguage = document.getElementById("benefitLanguage").checked;
            let simplifiedText = processText(emailText, useBenefitLanguage);
            Office.context.mailbox.item.body.setAsync(simplifiedText);
        }
    });
}

function processText(text, useBenefitLanguage) {
    let simplified = text.replace(/złożone wyrażenie/g, "prostsze słowo");
    if (useBenefitLanguage) {
        simplified += "\n\nTo rozwiązanie pomoże Ci...";
    }
    return simplified;
}