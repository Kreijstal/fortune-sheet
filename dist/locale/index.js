import en from "./en";
import zh from "./zh";
import es from "./es";
import hi from "./hi";
import zh_tw from "./zh_tw";
// @ts-ignore
var localeObj = { en: en, zh: zh, es: es, "zh-TW": zh_tw, hi: hi };
function locale(ctx) {
    var _a;
    var langsToTry = [ctx.lang || "", ((_a = ctx.lang) === null || _a === void 0 ? void 0 : _a.split("-")[0]) || ""];
    for (var i = 0; i < langsToTry.length; i += 1) {
        if (langsToTry[i] in localeObj) {
            return localeObj[langsToTry[i]];
        }
    }
    return localeObj.en;
}
export { locale };
