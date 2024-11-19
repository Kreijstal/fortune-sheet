import en from "./en";
import zh from "./zh";
import es from "./es";
import hi from "./hi";
import zh_tw from "./zh_tw";
// @ts-ignore
const localeObj = { en, zh, es, "zh-TW": zh_tw, hi };
function locale(ctx) {
    const langsToTry = [ctx.lang || "", ctx.lang?.split("-")[0] || ""];
    for (let i = 0; i < langsToTry.length; i += 1) {
        if (langsToTry[i] in localeObj) {
            return localeObj[langsToTry[i]];
        }
    }
    return localeObj.en;
}
export { locale };
