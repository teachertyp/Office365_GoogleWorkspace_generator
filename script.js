/* global i */

var lwords = [];
var words = [];
var emails = [];

function buttonclick() {
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
    var exportData = document.getElementById('OfficeImport').value = 'Ім’я користувача,Ім’я,Прізвище,Коротке ім’я,Посада,Відділ,Номер офісу,Робочий телефон,Мобільний телефон,Факс,Додаткова адреса електронної пошти,Адреса,Місто,Область або республіка,Поштовий індекс,Країна або регіон' + '\n' + ms365_import();
    //console.log(CyrUserFIO);
    var a = document.createElement('a');
    var orgua = orgua.replace('/', '-');
    var orgua = orgua.replace('/ /g', '_');
    var filename = orgua + 'data.csv';
    var data = 'data:application/csv;charset=utf-8,' + encodeURIComponent(exportData);
    a.href = data;
    a.target = '_blank';
    a.download = filename;
    document.body.appendChild(a);
    a.click();
}

function gsuite() {
    var CyrUserFIO = document.getElementById('Title').value;
    var domain = document.getElementById('yourdomain').value;
    var suffix = document.getElementById('emailsuffix').value;
    var orgua = document.getElementById('orgua').value;

    var orgua = orgua.replace('/', '-');
    var orgua = orgua.replace('/ /ig', '_');
    transliteration(CyrUserFIO);
    var exportData = document.getElementById('OfficeImport').value = 'First Name [Required],Last Name [Required],Email Address [Required],Password [Required],Password Hash Function [UPLOAD ONLY],Org Unit Path [Required],New Primary Email [UPLOAD ONLY],Recovery Email,Home Secondary Email,Work Secondary Email,Recovery Phone [MUST BE IN THE E.164 FORMAT],Work Phone,Home Phone,Mobile Phone,Work Address,Home Address,Employee ID,Employee Type,Employee Title,Manager Email,Department,Cost Center,Building ID,Floor Name,Floor Section,Change Password at Next Sign-In,New Status [UPLOAD ONLY],Advanced Protection Program enrollment' + '\n' + gsuite_import();
    //console.log(CyrUserFIO);
    var a = document.createElement('a');
    var filename = orgua + '-passwords.csv';
    var data = 'data:application/csv;charset=utf-8,' + encodeURIComponent(exportData);
    a.href = data;
    a.target = '_blank';
    a.download = filename;
    document.body.appendChild(a);
    a.click();
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

function ms365_import() {
    var str = "";
    var suffix = document.getElementById('emailsuffix').value;
    suffix = suffix != "" ? suffix + "_" : "";
    var domain = document.getElementById('yourdomain').value;
    var orgua = document.getElementById('orgua').value;
    for (i in words) {
        var fio = words[i].split('\t');
        console.log(fio);
        var enfio = lwords[i].split('\t');
        console.log(enfio);
        if (fio.length < 3) {
            fio = words[i].split(" ");
            enfio = lwords[i].split(" ");
            console.log(fio);
            console.log(enfio);
        }
        if (fio.length < 3) {
            fio[2] = 's';
            enfio[2] = 's';
        }
        if (Array.isArray(fio) && Array.isArray(enfio) && fio.length == 3 && enfio.length == 3) {
            console.log('Створюємо електронні адреси..........')
            var ns = 1;
            var sn = enfio[1][0];

            var email = suffix + enfio[0].toLowerCase() + "." + enfio[1][0].toLowerCase() + "." + enfio[2][0].toLowerCase() + "@" + domain;
            while (isset_email(email)) {
                sn = sn + enfio[1][ns];
                ns += 1;
                email = suffix + enfio[0].toLowerCase() + "." + sn.toLowerCase() + "." + enfio[2][0].toLowerCase() + "@" + domain;
            }
            emails[i] = email;
            str = str + email + ',' + fio[1] + ',' + fio[0]+ "," + fio[1] + fio[0] + ",," + orgua + ",,,,,,,,,,\n";
        }
    }
    console.log(emails);
    return str;
}

function gsuite_import() {
    var str = "";
    var suffix = document.getElementById('emailsuffix').value;
    suffix = suffix != "" ? suffix + "_" : "";
    var domain = document.getElementById('yourdomain').value;
    var orgua = document.getElementById('orgua').value;
    for (i in words) {
        var fio = words[i].split('\t');
        console.log(fio);
        var enfio = lwords[i].split('\t');
        console.log(enfio);
        if (fio.length < 3) {
            fio = words[i].split(" ");
            enfio = lwords[i].split(" ");
            console.log(fio);
            console.log(enfio);
        }
        if (fio.length < 3) {
            fio[2] = 's';
            enfio[2] = 's';
        }
        if (Array.isArray(fio) && Array.isArray(enfio) && fio.length == 3 && enfio.length == 3) {
            console.log('Створюємо електронні адреси..........')
            var ns = 1;
            var sn = enfio[1][0];

            var email = suffix + enfio[0].toLowerCase() + "." + enfio[1][0].toLowerCase() + "." + enfio[2][0].toLowerCase() + "@" + domain;
            while (isset_email(email)) {
                sn = sn + enfio[1][ns];
                ns += 1;
                email = suffix + enfio[0].toLowerCase() + "." + sn.toLowerCase() + "." + enfio[2][0].toLowerCase() + "@" + domain;
            }
            emails[i] = email;
            str = str + fio[1] + "," + fio[0] + "," + email + ", " + randomstring(8) + ",," + "/" + orgua + ",,,,,,,,,,,,,,,,,,,,,,\n";
        }
    }
    console.log(emails);
    return str;
}

function isset_email(e) {
    if (emails.includes(e)) {
        console.log(e + " вже є у списку");
        return true;
    }
    return false;
}

function transliteration(inputText) {
    var rules = [
        {'pattern': 'а', 'replace': 'a'},
        {'pattern': 'б', 'replace': 'b'},
        {'pattern': 'в', 'replace': 'v'},
        {'pattern': 'зг', 'replace': 'zgh'},
        {'pattern': 'Зг', 'replace': 'Zgh'},
        {'pattern': 'г', 'replace': 'h'},
        {'pattern': 'ґ', 'replace': 'g'},
        {'pattern': 'д', 'replace': 'd'},
        {'pattern': 'е', 'replace': 'e'},
        {'pattern': '^є', 'replace': 'ye'},
        {'pattern': 'є', 'replace': 'ie'},
        {'pattern': 'ж', 'replace': 'zh'},
        {'pattern': 'з', 'replace': 'z'},
        {'pattern': 'и', 'replace': 'y'},
        {'pattern': 'і', 'replace': 'i'},
        {'pattern': '^ї', 'replace': 'yi'},
        {'pattern': 'ї', 'replace': 'i'},
        {'pattern': '^й', 'replace': 'y'},
        {'pattern': 'й', 'replace': 'i'},
        {'pattern': 'к', 'replace': 'k'},
        {'pattern': 'л', 'replace': 'l'},
        {'pattern': 'м', 'replace': 'm'},
        {'pattern': 'н', 'replace': 'n'},
        {'pattern': 'о', 'replace': 'o'},
        {'pattern': 'п', 'replace': 'p'},
        {'pattern': 'р', 'replace': 'r'},
        {'pattern': 'с', 'replace': 's'},
        {'pattern': 'т', 'replace': 't'},
        {'pattern': 'у', 'replace': 'u'},
        {'pattern': 'ф', 'replace': 'f'},
        {'pattern': 'х', 'replace': 'kh'},
        {'pattern': 'ц', 'replace': 'ts'},
        {'pattern': 'ч', 'replace': 'ch'},
        {'pattern': 'ш', 'replace': 'sh'},
        {'pattern': 'щ', 'replace': 'shch'},
        {'pattern': 'ьо', 'replace': 'io'},
        {'pattern': 'ьї', 'replace': 'ii'},
        {'pattern': 'ь', 'replace': ''},
        {'pattern': '^ю', 'replace': 'yu'},
        {'pattern': 'ю', 'replace': 'iu'},
        {'pattern': '^я', 'replace': 'ya'},
        {'pattern': 'я', 'replace': 'ia'},
        {'pattern': 'А', 'replace': 'A'},
        {'pattern': 'Б', 'replace': 'B'},
        {'pattern': 'В', 'replace': 'V'},
        {'pattern': 'Г', 'replace': 'H'},
        {'pattern': 'Ґ', 'replace': 'G'},
        {'pattern': 'Д', 'replace': 'D'},
        {'pattern': 'Е', 'replace': 'E'},
        {'pattern': '^Є', 'replace': 'Ye'},
        {'pattern': 'Є', 'replace': 'Ie'},
        {'pattern': 'Ж', 'replace': 'Zh'},
        {'pattern': 'З', 'replace': 'Z'},
        {'pattern': 'И', 'replace': 'Y'},
        {'pattern': 'І', 'replace': 'I'},
        {'pattern': '^Ї', 'replace': 'Yi'},
        {'pattern': 'Ї', 'replace': 'I'},
        {'pattern': '^Й', 'replace': 'Y'},
        {'pattern': 'Й', 'replace': 'I'},
        {'pattern': 'К', 'replace': 'K'},
        {'pattern': 'Л', 'replace': 'L'},
        {'pattern': 'М', 'replace': 'M'},
        {'pattern': 'Н', 'replace': 'N'},
        {'pattern': 'О', 'replace': 'O'},
        {'pattern': 'П', 'replace': 'P'},
        {'pattern': 'Р', 'replace': 'R'},
        {'pattern': 'С', 'replace': 'S'},
        {'pattern': 'Т', 'replace': 'T'},
        {'pattern': 'У', 'replace': 'U'},
        {'pattern': 'Ф', 'replace': 'F'},
        {'pattern': 'Х', 'replace': 'Kh'},
        {'pattern': 'Ц', 'replace': 'Ts'},
        {'pattern': 'Ч', 'replace': 'Ch'},
        {'pattern': 'Ш', 'replace': 'Sh'},
        {'pattern': 'Щ', 'replace': 'Shch'},
        {'pattern': 'Ь', 'replace': ''},
        {'pattern': '^Ю', 'replace': 'Yu'},
        {'pattern': 'Ю', 'replace': 'Iu'},
        {'pattern': '^Я', 'replace': 'Ya'},
        {'pattern': 'Я', 'replace': 'Ia'},
        {'pattern': '’', 'replace': ''},
        {'pattern': '\'', 'replace': ''},
        {'pattern': '`', 'replace': ''}
    ];

    words = inputText.split(/[\n]/);
    var x = 0;
    for (var n in words) {


        var word = words[n];
        //console.log(word);
        cword = words[n];
        for (var ruleNumber in rules) {

            word = word.replace(
                    new RegExp(rules[ruleNumber]['pattern'], 'gm'),
                    rules[ruleNumber]['replace']
                    );
        }
        lwords[x] = word;
        x++;
        inputText = inputText.replace(words[n], word);
    }


    //console.log(inputText);
    console.log(words);
    return inputText; //.toUpperCase()
}
;
