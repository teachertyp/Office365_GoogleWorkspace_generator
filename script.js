/* global i */

var lwords = [];
var words = [];
var emails = [];
var rules = [
        { 'pattern': 'а', 'replace': 'a' },
        { 'pattern': 'б', 'replace': 'b' },
        { 'pattern': 'в', 'replace': 'v' },
        { 'pattern': 'зг', 'replace': 'zgh' },
        { 'pattern': 'Зг', 'replace': 'Zgh' },
        { 'pattern': 'г', 'replace': 'h' },
        { 'pattern': 'ґ', 'replace': 'g' },
        { 'pattern': 'д', 'replace': 'd' },
        { 'pattern': 'е', 'replace': 'e' },
        { 'pattern': '^є', 'replace': 'ye' },
        { 'pattern': 'є', 'replace': 'ie' },
        { 'pattern': 'ж', 'replace': 'zh' },
        { 'pattern': 'з', 'replace': 'z' },
        { 'pattern': 'и', 'replace': 'y' },
        { 'pattern': 'і', 'replace': 'i' },
        { 'pattern': '^ї', 'replace': 'yi' },
        { 'pattern': 'ї', 'replace': 'i' },
        { 'pattern': '^й', 'replace': 'y' },
        { 'pattern': 'й', 'replace': 'i' },
        { 'pattern': 'к', 'replace': 'k' },
        { 'pattern': 'л', 'replace': 'l' },
        { 'pattern': 'м', 'replace': 'm' },
        { 'pattern': 'н', 'replace': 'n' },
        { 'pattern': 'о', 'replace': 'o' },
        { 'pattern': 'п', 'replace': 'p' },
        { 'pattern': 'р', 'replace': 'r' },
        { 'pattern': 'с', 'replace': 's' },
        { 'pattern': 'т', 'replace': 't' },
        { 'pattern': 'у', 'replace': 'u' },
        { 'pattern': 'ф', 'replace': 'f' },
        { 'pattern': 'х', 'replace': 'kh' },
        { 'pattern': 'ц', 'replace': 'ts' },
        { 'pattern': 'ч', 'replace': 'ch' },
        { 'pattern': 'ш', 'replace': 'sh' },
        { 'pattern': 'щ', 'replace': 'shch' },
        { 'pattern': 'ьо', 'replace': 'io' },
        { 'pattern': 'ьї', 'replace': 'ii' },
        { 'pattern': 'ь', 'replace': '' },
        { 'pattern': '^ю', 'replace': 'yu' },
        { 'pattern': 'ю', 'replace': 'iu' },
        { 'pattern': '^я', 'replace': 'ya' },
        { 'pattern': 'я', 'replace': 'ia' },
        { 'pattern': 'А', 'replace': 'A' },
        { 'pattern': 'Б', 'replace': 'B' },
        { 'pattern': 'В', 'replace': 'V' },
        { 'pattern': 'Г', 'replace': 'H' },
        { 'pattern': 'Ґ', 'replace': 'G' },
        { 'pattern': 'Д', 'replace': 'D' },
        { 'pattern': 'Е', 'replace': 'E' },
        { 'pattern': '^Є', 'replace': 'Ye' },
        { 'pattern': 'Є', 'replace': 'Ie' },
        { 'pattern': 'Ж', 'replace': 'Zh' },
        { 'pattern': 'З', 'replace': 'Z' },
        { 'pattern': 'И', 'replace': 'Y' },
        { 'pattern': 'І', 'replace': 'I' },
        { 'pattern': '^Ї', 'replace': 'Yi' },
        { 'pattern': 'Ї', 'replace': 'I' },
        { 'pattern': '^Й', 'replace': 'Y' },
        { 'pattern': 'Й', 'replace': 'I' },
        { 'pattern': 'К', 'replace': 'K' },
        { 'pattern': 'Л', 'replace': 'L' },
        { 'pattern': 'М', 'replace': 'M' },
        { 'pattern': 'Н', 'replace': 'N' },
        { 'pattern': 'О', 'replace': 'O' },
        { 'pattern': 'П', 'replace': 'P' },
        { 'pattern': 'Р', 'replace': 'R' },
        { 'pattern': 'С', 'replace': 'S' },
        { 'pattern': 'Т', 'replace': 'T' },
        { 'pattern': 'У', 'replace': 'U' },
        { 'pattern': 'Ф', 'replace': 'F' },
        { 'pattern': 'Х', 'replace': 'Kh' },
        { 'pattern': 'Ц', 'replace': 'Ts' },
        { 'pattern': 'Ч', 'replace': 'Ch' },
        { 'pattern': 'Ш', 'replace': 'Sh' },
        { 'pattern': 'Щ', 'replace': 'Shch' },
        { 'pattern': 'Ь', 'replace': '' },
        { 'pattern': '^Ю', 'replace': 'Yu' },
        { 'pattern': 'Ю', 'replace': 'Iu' },
        { 'pattern': '^Я', 'replace': 'Ya' },
        { 'pattern': 'Я', 'replace': 'Ia' },
        { 'pattern': '’', 'replace': '' },
        { 'pattern': '\'', 'replace': '' },
        { 'pattern': '`', 'replace': '' }
    ];
    var lwords = [];
    var words = [];
    var emails = [];
    
    function buttonclick() {
        emails = [];
        if (document.getElementById("ms365").checked) {
            ms365();
        } else if (document.getElementById("gsuite").checked) {
            gsuite();
        }
    }
    
    function ms365() {
        var CyrUserFIO = document.getElementById('Title').value;
        var domain = document.getElementById('yourdomain').value;
        var suffix = document.getElementById('emailsuffix').value;
        var orgua = document.getElementById('orgua').value;
        transliteration(CyrUserFIO);
        var exportData = 'Ім’я користувача,Ім’я,Прізвище,Коротке ім’я,Посада,Відділ,Номер офісу,Робочий телефон,Мобільний телефон,Факс,Додаткова адреса електронної пошти,Адреса,Місто,Область або республіка,Поштовий індекс,Країна або регіон\n' + ms365_import();
        downloadCSV(exportData, orgua + 'data.csv');
    }
    
    function gsuite() {
        var CyrUserFIO = document.getElementById('Title').value;
        var domain = document.getElementById('yourdomain').value;
        var suffix = document.getElementById('emailsuffix').value;
        var orgua = document.getElementById('orgua').value;
        transliteration(CyrUserFIO);
        var exportData = 'First Name [Required],Last Name [Required],Email Address [Required],Password [Required],Password Hash Function [UPLOAD ONLY],Org Unit Path [Required],New Primary Email [UPLOAD ONLY],Recovery Email,Home Secondary Email,Work Secondary Email,Recovery Phone [MUST BE IN THE E.164 FORMAT],Work Phone,Home Phone,Mobile Phone,Work Address,Home Address,Employee ID,Employee Type,Employee Title,Manager Email,Department,Cost Center,Building ID,Floor Name,Floor Section,Change Password at Next Sign-In,New Status [UPLOAD ONLY],Advanced Protection Program enrollment\n' + gsuite_import();
        downloadCSV(exportData, orgua + '-passwords.csv');
    }
    
    function downloadCSV(data, filename) {
        document.getElementById('OfficeImport').value = data;
        var a = document.createElement('a');
        var dataUri = 'data:application/csv;charset=utf-8,' + encodeURIComponent(data);
        a.href = dataUri;
        a.target = '_blank';
        a.download = filename;
        document.body.appendChild(a);
        a.click();
    }
    
    function importData(rules, suffix, domain, orgua, importFunction) {
        var str = "";
        for (var i = 0; i < words.length; i++) {
            var fio = words[i].split('\t');
            var enfio = lwords[i].split('\t');
    
            if (fio.length < 3) {
                fio = words[i].split(" ");
                enfio = lwords[i].split(" ");
            }
            
            if (fio.length < 3) {
                fio[2] = 's';
                enfio[2] = 's';
            }
    
            if (Array.isArray(fio) && Array.isArray(enfio) && fio.length === 3 && enfio.length === 3) {
                var ns = 1;
                var sn = enfio[1][0];
                var email = suffix + enfio[0].toLowerCase() + "." + enfio[1][0].toLowerCase() + "." + enfio[2][0].toLowerCase() + "@" + domain;
    
                while (emails.includes(email)) {
                    sn = sn + enfio[1][ns];
                    ns += 1;
                    email = suffix + enfio[0].toLowerCase() + "." + sn.toLowerCase() + "." + enfio[2][0].toLowerCase() + "@" + domain;
                }
    
                emails[i] = email;
                str += importFunction(fio, enfio, email, orgua);
            }
        }
        return str;
    }
    
    function ms365_import() {
        var suffix = document.getElementById('emailsuffix').value;
        suffix = suffix ? suffix + "_" : "";
        var domain = document.getElementById('yourdomain').value;
        var orgua = document.getElementById('orgua').value;
        return importData(rules, suffix, domain, orgua, ms365_importLine);
    }
    
    function ms365_importLine(fio, enfio, email, orgua) {
        return email + ',' + fio[1] + ',' + fio[0] + "," + fio[1] + fio[0] + ",," + orgua + ",,,,,,,,,,\n";
    }
    
    function gsuite_import() {
        var suffix = document.getElementById('emailsuffix').value;
        suffix = suffix ? suffix + "_" : "";
        var domain = document.getElementById('yourdomain').value;
        var orgua = document.getElementById('orgua').value;
        return importData(rules, suffix, domain, orgua, gsuite_importLine);
    }
    
    function gsuite_importLine(fio, enfio, email, orgua) {
        return fio[1] + "," + fio[0] + "," + email + ", " + randomstring(8) + ",," + "/" + orgua + ",,,,,,,,,,,,,,,,,,,,,,\n";
    }
    
    function getRandomInt(max) {
        return Math.floor(Math.random() * Math.floor(max));
    }
    
    function randomstring(L) {
        var s = '';
        var randomchar = function () {
            var n = Math.floor(Math.random() * 62);
            if (n < 10)
                return n; //1-10
            if (n < 36)
                return String.fromCharCode(n + 55); //A-Z
            return String.fromCharCode(n + 61); //a-z
        }
        while (s.length < L)
            s += randomchar();
        return s;
    }
    
function transliteration(inputText) {
    words = inputText.split(/[\n]/);
    for (var n = 0; n < words.length; n++) {
        var word = words[n].split(/[ \t]+/);
        var cword = [];
        var cyrwords = [];

        for (var k = 0; k < word.length; k++) {
            var lattext = word[k];
            cyrwords.push(word[k]);

            for (var ruleNumber = 0; ruleNumber < rules.length; ruleNumber++) {
                lattext = lattext.replace(
                    new RegExp(rules[ruleNumber]['pattern'], 'gm'),
                    rules[ruleNumber]['replace']
                );
            }

            cword.push(lattext);
        }

        words[n] = cyrwords.join(' ');
        lwords[n] = cword.join(' ');
    }
}
