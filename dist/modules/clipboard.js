var clipboard = /** @class */ (function () {
    function clipboard() {
    }
    clipboard.writeHtml = function (str) {
        var _a;
        try {
            var ele_1 = document.getElementById("fortune-copy-content");
            if (!ele_1) {
                ele_1 = document.createElement("div");
                ele_1.setAttribute("contentEditable", "true");
                ele_1.id = "fortune-copy-content";
                ele_1.style.position = "fixed";
                ele_1.style.height = "0";
                ele_1.style.width = "0";
                ele_1.style.left = "-10000px";
                (_a = document.querySelector(".fortune-container")) === null || _a === void 0 ? void 0 : _a.append(ele_1);
            }
            var previouslyFocusedElement_1 = document.activeElement;
            ele_1.style.display = "block";
            ele_1.innerHTML = str;
            ele_1.focus({ preventScroll: true });
            document.execCommand("selectAll");
            document.execCommand("copy");
            setTimeout(function () {
                var _a;
                ele_1 === null || ele_1 === void 0 ? void 0 : ele_1.blur();
                (_a = previouslyFocusedElement_1 === null || previouslyFocusedElement_1 === void 0 ? void 0 : previouslyFocusedElement_1.focus) === null || _a === void 0 ? void 0 : _a.call(previouslyFocusedElement_1);
            }, 10);
        }
        catch (e) {
            console.error(e);
        }
    };
    return clipboard;
}());
export default clipboard;
