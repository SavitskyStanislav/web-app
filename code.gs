// Константы приложения
const CONFIG = {
  API_KEY: "ed5af7fdee3e1fdfc79d1e688160a043f43f8841",
  SHEET_ID: '1IFIlcDsSR0dMENaodW8ZUPfjbxBmwec2c7qH_rAgEXo',
  DADATA_API: {
    URL: "https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party",
    HEADERS: function() {
      return {
        "Authorization": "Token " + this.API_KEY,
        "Content-Type": "application/json",
        "Accept": "application/json"
      };
    }
  }
};

// Класс для работы с документами
class DocumentHandler {
  // Функция создания документа
  static createDocument(data, documentId, folderId, docVisibleName) {
    try {
      const file = DriveApp.getFileById(documentId);
      const folder = DriveApp.getFolderById(folderId);
      const copiedFile = file.makeCopy(folder);
      
      // Устанавливаем имя
      const newFileName = docVisibleName;
      copiedFile.setName(newFileName);
      
      // Обрабатываем документ
      const doc = DocumentApp.openById(copiedFile.getId());
      const body = doc.getBody();

      // Заменяем плейсхолдеры
      Object.entries(data).forEach(([placeholder, replacement]) => {
        body.replaceText(placeholder, replacement);
      });

      doc.saveAndClose();

      // Настройка прав доступа
      copiedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      const docId = copiedFile.getId();
      return {
        docxUrl: `https://docs.google.com/document/d/${docId}/export?format=docx`,
        pdfUrl: `https://docs.google.com/document/d/${docId}/export?format=pdf`
      };
    } catch (error) {
      Logger.log(`Ошибка при создании документа: ${error.message}`);
      throw new Error(`Ошибка при создании документа: ${error.message}`);
    }
  }
}

// Класс для работы с данными компании
class CompanyDataService {
  static async getCompanyData(inn) {
    if (!AuthorizationService.checkAuthorization()) {
      throw new Error("Нет необходимых разрешений");
    }

    // Validate INN format
    if (!this._validateINN(inn)) {
      throw new Error("Неверный формат ИНН");
    }

    const companyData = await this._fetchCompanyDataFromDaData(inn);
    return this._formatCompanyData(companyData);
  }

  static _validateINN(inn) {
    // Basic INN validation (10 or 12 digits)
    return /^\d{10}(\d{2})?$/.test(inn);
  }

  static async _fetchCompanyDataFromDaData(inn) {
    // Validate API configuration
    if (!CONFIG.API_KEY) {
      throw new Error("API ключ не настроен");
    }

    const options = {
      method: "post",
      payload: JSON.stringify({ query: inn }),
      headers: {
        "Authorization": `Token ${CONFIG.API_KEY}`,
        "Content-Type": "application/json",
        "Accept": "application/json"
      },
      muteHttpExceptions: true,
      validateHttpsCertificates: true
    };

    try {
      Logger.log(`Отправка запроса к DaData для ИНН: ${inn}`);
      const response = await UrlFetchApp.fetch(CONFIG.DADATA_API.URL, options);
      const responseCode = response.getResponseCode();
      const contentText = response.getContentText();

      Logger.log(`Получен ответ с кодом: ${responseCode}`);

      if (responseCode === 403) {
        throw new Error("Ошибка авторизации в DaData API. Проверьте правильность API ключа");
      }

      if (responseCode !== 200) {
        throw new Error(`Ошибка DaData API: ${responseCode} - ${contentText}`);
      }

      const json = JSON.parse(contentText);
      
      if (!json.suggestions?.[0]) {
        throw new Error(`Компания с ИНН ${inn} не найдена`);
      }

      return json.suggestions[0].data;

    } catch (error) {
      Logger.log(`Ошибка при запросе к DaData: ${error.message}`);
      
      // Add diagnostic information
      if (error.message.includes('DNS lookup failed')) {
        throw new Error("Не удалось подключиться к серверу DaData. Проверьте доступность сервиса и сетевое подключение");
      }
      
      throw error;
    }
  }

  static _formatCompanyData(company) {
    // Add null checks and default values
    return {
      type: company?.type || "",
      name: TextFormatter.formatOrgForm(company?.name?.full_with_opf || ""),
      shortName: company?.name?.short_with_opf || "",
      shortNameWithoutQuotes: TextFormatter.removeQuotes(company?.name?.short_with_opf || ""),
      inn: company?.inn || "",
      ogrn: company?.ogrn || "",
      kpp: company?.kpp || "",
      address: company?.address?.data?.source || "",
      managementName: company?.management?.name || "",
      managementPost: TextFormatter.formatPosition(company?.management?.post || ""),
      managementPostGc: TextFormatter.declinePosition(company?.management?.post || ""),
      managementNameGc: TextFormatter.declineNameInGc(company?.management?.name || ""),
      managementNameInitials: TextFormatter.getInitials(company?.management?.name || "")
    };
  }

  // Helper method to test API connectivity
  static async testAPIConnection() {
    try {
      const response = await UrlFetchApp.fetch(CONFIG.DADATA_API.URL, {
        method: "get",
        headers: {
          "Authorization": `Token ${CONFIG.API_KEY}`
        },
        muteHttpExceptions: true
      });
      
      const responseCode = response.getResponseCode();
      Logger.log(`API test connection response code: ${responseCode}`);
      
      return {
        success: responseCode !== 403,
        statusCode: responseCode,
        message: responseCode === 403 ? "Invalid API key" : "Connection successful"
      };
    } catch (error) {
      Logger.log(`API test connection error: ${error.message}`);
      return {
        success: false,
        statusCode: 500,
        message: error.message
      };
    }
  }
}

// Класс для работы с таблицами
class SpreadsheetService {
  static getParticipationData() {
    try {
      const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getActiveSheet();
      const data = sheet.getDataRange().getValues();
      
      // Пропускаем заголовок и формируем объект
      return data.slice(1).reduce((acc, [format, options]) => {
        if (format && options) {
          acc[format] = options;
        }
        return acc;
      }, {});
    } catch (error) {
      Logger.log(`Ошибка при чтении таблицы: ${error.message}`);
      throw new Error(`Ошибка при чтении таблицы: ${error.message}`);
    }
  }
}

// Класс для форматирования текста
class TextFormatter {
  static formatPosition(position) {
    return position.charAt(0).toUpperCase() + position.slice(1).toLowerCase();
  }

  static declinePosition(position) {
    const positionLower = position.toLowerCase();
    if (positionLower.includes("генеральный директор")) {
      return "Генерального директора";
    }
    if (positionLower.includes("директор")) {
      return "Директора";
    }
    return position;
  }

  static _splitFullName(fullName) {
    const parts = fullName.split(' ');
    if (parts.length < 3) return fullName;

    const person = {
      first: parts[1],
      middle: parts[2],
      last: parts[0]
    };

    return person
  }

  static declineNameInGc(fullName) {
    const personData = TextFormatter._splitFullName(fullName);
    const genetiveNameResult = TextFormatter.declensionOfFullName(personData, 'genitive');
    return `${genetiveNameResult.last} ${genetiveNameResult.first} ${genetiveNameResult.middle}`;
  }

  static getInitials(fullName) {
    const parts = fullName.split(' ');
    if (parts.length < 3) return fullName;

    return `${parts[0]} ${parts[1].charAt(0)}.${parts[2].charAt(0)}.`;
  }

  static formatOrgForm(companyName) {
    const orgFormMatch = companyName.match(/^[^"]+(?=\s*")/);
    if (!orgFormMatch) return companyName;

    const orgForm = orgFormMatch[0].trim();
    const formattedOrgForm = orgForm.charAt(0).toUpperCase() + orgForm.slice(1).toLowerCase();
    return companyName.replace(orgForm, formattedOrgForm);
  }

  static removeQuotes(text) {
    return text.replace(/["']/g, '');
  }

  static declensionOfFullName(person, gcase) {
    // Предопределенные значения вынесены внутрь функции
    const predef = {
      genders: ['male', 'female', 'androgynous'],
      nametypes: ['first', 'middle', 'last'],
      cases: ['nominative', 'genitive', 'dative', 'accusative', 'instrumental', 'prepositional']
    };

    // Правила склонения
    const rules = { "lastname": { "exceptions": [{ "gender": "androgynous", "test": ["бонч", "абдул", "белиц", "гасан", "дюссар", "дюмон", "книппер", "корвин", "ван", "шолом", "тер", "призван", "мелик", "вар", "фон"], "mods": [".", ".", ".", ".", "."], "tags": ["first_word"] }, { "gender": "androgynous", "test": ["дюма", "тома", "дега", "люка", "ферма", "гамарра", "петипа", "шандра", "скаля", "каруана"], "mods": [".", ".", ".", ".", "."] }, { "gender": "androgynous", "test": ["гусь", "ремень", "камень", "онук", "богода", "нечипас", "долгопалец", "маненок", "рева", "кива"], "mods": [".", ".", ".", ".", "."] }, { "gender": "androgynous", "test": ["вий", "сой", "цой", "хой"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "androgynous", "test": ["я"], "mods": [".", ".", ".", ".", "."] }], "suffixes": [{ "gender": "female", "test": ["б", "в", "г", "д", "ж", "з", "й", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ь"], "mods": [".", ".", ".", ".", "."] }, { "gender": "androgynous", "test": ["гава", "орота"], "mods": [".", ".", ".", ".", "."] }, { "gender": "female", "test": ["ска", "цка"], "mods": ["-ой", "-ой", "-ую", "-ой", "-ой"] }, { "gender": "female", "test": ["цкая", "ская", "ная", "ая"], "mods": ["--ой", "--ой", "--ую", "--ой", "--ой"] }, { "gender": "female", "test": ["яя"], "mods": ["--ей", "--ей", "--юю", "--ей", "--ей"] }, { "gender": "female", "test": ["на"], "mods": ["-ой", "-ой", "-у", "-ой", "-ой"] }, { "gender": "male", "test": ["иной"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "male", "test": ["уй"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "androgynous", "test": ["ца"], "mods": ["-ы", "-е", "-у", "-ей", "-е"] }, { "gender": "male", "test": ["рих"], "mods": ["а", "у", "а", "ом", "е"] }, { "gender": "androgynous", "test": ["ия"], "mods": [".", ".", ".", ".", "."] }, { "gender": "androgynous", "test": ["иа", "аа", "оа", "уа", "ыа", "еа", "юа", "эа"], "mods": [".", ".", ".", ".", "."] }, { "gender": "male", "test": ["их", "ых"], "mods": [".", ".", ".", ".", "."] }, { "gender": "androgynous", "test": ["о", "е", "э", "и", "ы", "у", "ю"], "mods": [".", ".", ".", ".", "."] }, { "gender": "female", "test": ["ова", "ева"], "mods": ["-ой", "-ой", "-у", "-ой", "-ой"] }, { "gender": "androgynous", "test": ["га", "ка", "ха", "ча", "ща", "жа", "ша"], "mods": ["-и", "-е", "-у", "-ой", "-е"] }, { "gender": "androgynous", "test": ["а"], "mods": ["-ы", "-е", "-у", "-ой", "-е"] }, { "gender": "male", "test": ["ь"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "androgynous", "test": ["ия"], "mods": ["-и", "-и", "-ю", "-ей", "-и"] }, { "gender": "androgynous", "test": ["я"], "mods": ["-и", "-е", "-ю", "-ей", "-е"] }, { "gender": "male", "test": ["ей"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "male", "test": ["ян", "ан", "йн"], "mods": ["а", "у", "а", "ом", "е"] }, { "gender": "male", "test": ["ынец", "обец"], "mods": ["--ца", "--цу", "--ца", "--цем", "--це"] }, { "gender": "male", "test": ["онец", "овец"], "mods": ["--ца", "--цу", "--ца", "--цом", "--це"] }, { "gender": "male", "test": ["ай"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "male", "test": ["кой"], "mods": ["-го", "-му", "-го", "--им", "-м"] }, { "gender": "male", "test": ["гой"], "mods": ["-го", "-му", "-го", "--им", "-м"] }, { "gender": "male", "test": ["ой"], "mods": ["-го", "-му", "-го", "--ым", "-м"] }, { "gender": "male", "test": ["ах", "ив"], "mods": ["а", "у", "а", "ом", "е"] }, { "gender": "male", "test": ["ший", "щий", "жий", "ний"], "mods": ["--его", "--ему", "--его", "-м", "--ем"] }, { "gender": "male", "test": ["ый"], "mods": ["--ого", "--ому", "--ого", "-м", "--ом"] }, { "gender": "male", "test": ["кий"], "mods": ["--ого", "--ому", "--ого", "-м", "--ом"] }, { "gender": "male", "test": ["ий"], "mods": ["-я", "-ю", "-я", "-ем", "-и"] }, { "gender": "male", "test": ["ок"], "mods": ["--ка", "--ку", "--ка", "--ком", "--ке"] }, { "gender": "male", "test": ["ец"], "mods": ["--ца", "--цу", "--ца", "--цом", "--це"] }, { "gender": "male", "test": ["ц", "ч", "ш", "щ"], "mods": ["а", "у", "а", "ем", "е"] }, { "gender": "male", "test": ["ен", "нн", "он", "ун"], "mods": ["а", "у", "а", "ом", "е"] }, { "gender": "male", "test": ["в", "н"], "mods": ["а", "у", "а", "ым", "е"] }, { "gender": "male", "test": ["б", "г", "д", "ж", "з", "к", "л", "м", "п", "р", "с", "т", "ф", "х"], "mods": ["а", "у", "а", "ом", "е"] }] }, "firstname": { "exceptions": [{ "gender": "male", "test": ["лев"], "mods": ["--ьва", "--ьву", "--ьва", "--ьвом", "--ьве"] }, { "gender": "male", "test": ["пётр"], "mods": ["---етра", "---етру", "---етра", "---етром", "---етре"] }, { "gender": "male", "test": ["павел"], "mods": ["--ла", "--лу", "--ла", "--лом", "--ле"] }, { "gender": "male", "test": ["яша"], "mods": ["-и", "-е", "-у", "-ей", "-е"] }, { "gender": "male", "test": ["шота"], "mods": [".", ".", ".", ".", "."] }, { "gender": "female", "test": ["рашель", "нинель", "николь", "габриэль", "даниэль"], "mods": [".", ".", ".", ".", "."] }], "suffixes": [{ "gender": "androgynous", "test": ["е", "ё", "и", "о", "у", "ы", "э", "ю"], "mods": [".", ".", ".", ".", "."] }, { "gender": "female", "test": ["б", "в", "г", "д", "ж", "з", "й", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "х", "ц", "ч", "ш", "щ", "ъ"], "mods": [".", ".", ".", ".", "."] }, { "gender": "female", "test": ["ь"], "mods": ["-и", "-и", ".", "ю", "-и"] }, { "gender": "male", "test": ["ь"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "androgynous", "test": ["га", "ка", "ха", "ча", "ща", "жа"], "mods": ["-и", "-е", "-у", "-ой", "-е"] }, { "gender": "female", "test": ["ша"], "mods": ["-и", "-е", "-у", "-ей", "-е"] }, { "gender": "androgynous", "test": ["а"], "mods": ["-ы", "-е", "-у", "-ой", "-е"] }, { "gender": "female", "test": ["ия"], "mods": ["-и", "-и", "-ю", "-ей", "-и"] }, { "gender": "female", "test": ["а"], "mods": ["-ы", "-е", "-у", "-ой", "-е"] }, { "gender": "female", "test": ["я"], "mods": ["-и", "-е", "-ю", "-ей", "-е"] }, { "gender": "male", "test": ["ия"], "mods": ["-и", "-и", "-ю", "-ей", "-и"] }, { "gender": "male", "test": ["я"], "mods": ["-и", "-е", "-ю", "-ей", "-е"] }, { "gender": "male", "test": ["ей"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "male", "test": ["ий"], "mods": ["-я", "-ю", "-я", "-ем", "-и"] }, { "gender": "male", "test": ["й"], "mods": ["-я", "-ю", "-я", "-ем", "-е"] }, { "gender": "male", "test": ["б", "в", "г", "д", "ж", "з", "к", "л", "м", "н", "п", "р", "с", "т", "ф", "х", "ц", "ч"], "mods": ["а", "у", "а", "ом", "е"] }, { "gender": "androgynous", "test": ["ния", "рия", "вия"], "mods": ["-и", "-и", "-ю", "-ем", "-ем"] }] }, "middlename": { "suffixes": [{ "gender": "male", "test": ["ич"], "mods": ["а", "у", "а", "ем", "е"] }, { "gender": "female", "test": ["на"], "mods": ["-ы", "-е", "-у", "-ой", "-е"] }] } };

    // Вспомогательная функция для проверки наличия элемента в массиве
    const contains = (arr, x) => arr.includes(x);

    // Функция для определения пола по отчеству
    const detectGender = (middle) => {
      const ending = middle.toLowerCase().slice(-2);
      if (ending === 'ич') return 'male';
      if (ending === 'на') return 'female';
      return 'androgynous';
    };

    // Функция для поиска правил в исключениях или суффиксах
    const findRuleLocal = (gender, name, ruleset, matchWholeWord, tags) => {
      for (const rule of ruleset) {
        if (rule.tags) {
          const commonTags = rule.tags.filter(tag => !tags.includes(tag));
          if (commonTags.length === 0) continue;
        }
        
        if (rule.gender !== 'androgynous' && gender !== rule.gender) continue;

        const normalizedName = name.toLowerCase();
        for (const sample of rule.test) {
          const test = matchWholeWord ? normalizedName : 
            normalizedName.slice(-sample.length);
          if (test === sample) return rule;
        }
      }
      return false;
    };

    // Функция для поиска групп правил
    const findRuleGlobal = (gender, name, nametypeRulesets, features = {}) => {
      const tags = Object.keys(features);
      
      if (nametypeRulesets.exceptions) {
        const rule = findRuleLocal(gender, name, nametypeRulesets.exceptions, true, tags);
        if (rule) return rule;
      }
      
      return findRuleLocal(gender, name, nametypeRulesets.suffixes, false, tags);
    };

    // Функция применения найденного правила
    const applyRule = (name, gcase, rule) => {
      let mod = gcase === 'nominative' ? '.' : rule.mods[predef.cases.indexOf(gcase) - 1];

      return [...mod].reduce((acc, chr) => {
        switch (chr) {
          case '.': return acc;
          case '-': return acc.slice(0, -1);
          default: return acc + chr;
        }
      }, name);
    };

    // Функция склонения
    const inflect = (gender, name, gcase, nametype) => {
      const nametypeRulesets = rules[nametype];
      const parts = name.split('-');
      
      return parts.map((part, i) => {
        const firstWord = i === 0 && parts.length > 1;
        const rule = findRuleGlobal(gender, part, nametypeRulesets, { first_word: firstWord });
        return rule ? applyRule(part, gcase, rule) : part;
      }).join('-');
    };

    // Основная логика
    const result = {};

    // Определение пола
    if (person.gender != null) {
      if (!contains(predef.genders, person.gender)) {
        throw new Error(`Invalid gender: ${person.gender}`);
      }
      result.gender = person.gender;
    } else if (person.middle != null) {
      result.gender = detectGender(person.middle);
    } else {
      throw new Error('Unknown gender');
    }

    // Проверка падежа
    if (!contains(predef.cases, gcase)) {
      throw new Error(`Invalid case: ${gcase}`);
    }

    // Склонение всех частей имени
    for (const nametype of predef.nametypes) {
      if (person[nametype] != null) {
        result[nametype] = inflect(result.gender, person[nametype], gcase, nametype + 'name');
      }
    }

    return result;
  };
}

// Класс для проверки авторизации и разрешений
class AuthorizationService {
  static checkAuthorization() {
    return true; // TODO: Реализовать реальную проверку авторизации
  }

  static testAllPermissions() {
    Logger.log("Начало проверки разрешений");
    
    try {
      // Проверка Drive
      DriveApp.getFiles();
      Logger.log("✓ Доступ к Drive успешен");
      
      // Проверка Documents
      const testDoc = DocumentApp.create("Test Document");
      testDoc.removeFromTrash();
      Logger.log("✓ Доступ к Documents успешен");

      // Проверка Spreadsheets
      const testSheet = SpreadsheetApp.create("Test Document");
      testSheet.removeFromTrash();
      Logger.log("✓ Доступ к Spreadsheets успешен");
      
      // Проверка UrlFetchApp
      UrlFetchApp.fetch("https://www.google.com");
      Logger.log("✓ Внешние запросы работают");
      
      Logger.log("Все разрешения работают корректно");
      return true;
    } catch (error) {
      Logger.log(`Ошибка при проверке разрешений: ${error.message}`);
      return false;
    }
  }
}

// Основные точки входа в приложение
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Конструктор документов v3.1.0');
}

function replacePlaceholdersInDocument(data, documentId, folderId, docVisibleName) {
  return DocumentHandler.createDocument(data, documentId, folderId, docVisibleName);
}

function getCompanyDataByInn(companyInn) {
  return CompanyDataService.getCompanyData(companyInn);
}

function getParticipationData() {
  return SpreadsheetService.getParticipationData();
}

