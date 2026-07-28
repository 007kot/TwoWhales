document.addEventListener('DOMContentLoaded', function() {
    const contractSelect = document.getElementById('contract');
    const propertySelect = document.getElementById('property_type');
    const initialPaymentInput = document.getElementById('initial_payment');
    const installmentPeriodInput = document.getElementById('installment_period');
    const birthdateInput = document.getElementById('birthdate');
    const passportIssueDate = document.getElementById('passport_issue_date');

    const apartmentData = document.getElementById('apartment_data');
    const parkingData = document.getElementById('parking_data');
    const roomsField = document.getElementById('rooms_field');
    const roomsInput = document.getElementById('rooms');
    const totalArea = document.getElementById('total_area');
    const installmentCheckbox = document.getElementById('supportCheckbox');

    function setDateLimits() {
        const today = new Date();
        const maxDate = new Date(today); // Сегодняшняя дата
        const minDate = new Date(today);
        minDate.setFullYear(today.getFullYear() - 100); // 100 лет назад
        
        const formatDate = (date) => date.toISOString().split('T')[0];
        
        birthdateInput.max = formatDate(maxDate);
        birthdateInput.min = formatDate(minDate);
        passportIssueDate.max = formatDate(maxDate);
        passportIssueDate.min = formatDate(minDate);
    }

    function updatePaymentFields() {
        if (contractSelect.value === 'installment') {
            // Для "рассрочки" - включаем поля
            initialPaymentInput.disabled = false;
            installmentPeriodInput.disabled = false;
            installmentCheckbox.disabled = false;
        }else {
            // Для "налички" и другого - дизейблим и очищаем поля
            initialPaymentInput.disabled = true;
            installmentPeriodInput.disabled = true;
            initialPaymentInput.value = '';
            installmentPeriodInput.value = '';
            installmentCheckbox.disabled = true;
        }
    }

    function updatePropertyFields() {
        const propertyType = propertySelect.value;
        const isParking = propertyType === 'parking';
        // Коммерция использует тот же набор полей, что и квартира, но без комнат
        const isCommerce = propertyType === 'commerce';

        apartmentData.style.display = isParking ? 'none' : 'block';
        parkingData.style.display = isParking ? 'block' : 'none';

        roomsField.style.display = isCommerce ? 'none' : 'block';
        if (isCommerce) {
            roomsInput.value = '';
        }

        if (isParking) {
            totalArea.value = 25;
            totalArea.disabled = true;
        } else {
            // Чистим площадь только если уходим с парковки,
            // чтобы переключение квартира <-> коммерция не стирало ввод
            if (totalArea.disabled) {
                totalArea.value = '';
            }
            totalArea.disabled = false;
        }
    }

    setDateLimits();
    updatePropertyFields();
    updatePaymentFields();

    contractSelect.addEventListener('change', updatePaymentFields);
    propertySelect.addEventListener('change', updatePropertyFields);

    const submitBtn = document.getElementById('submit');
    const graphBtn = document.getElementById('graph');

    submitBtn.addEventListener('click', function(event) {
        event.preventDefault(); 

        const form = document.querySelector('form');

        if (form.checkValidity()) {
            generateContract();
        } else {
            form.reportValidity();
        }
    });

    graphBtn.addEventListener('click', function(event) {
        event.preventDefault(); 
        displayPaymentPreview();
    });

    Inputmask({
        mask: '+7 (999) 999-99-99',
        placeholder: '_',
        showMaskOnHover: true,
        showMaskOnFocus: true
    }).mask(document.getElementById('phone_number'));
});

async function generateContract() {
    try {
        // 1. Собираем данные из формы
        const formData = collectFormData();

        // 2. Загружаем и заполняем DOCX шаблон
        const docxBuffer = await fillDocxTemplate(formData);

        // 3. Конвертируем DOCX в PDF
        const pdfBlob = await convertDocxToPdf(docxBuffer);

        // 4. Скачиваем PDF
        downloadFile(pdfBlob, 'договор.docx');

    } catch (error) {
        console.error('Ошибка:', error);
        alert('Произошла ошибка при генерации документа');
    }
}

// Сбор данных из формы
function collectFormData() {
    const value = (id) => document.getElementById(id).value;

    const total_area = Number(value('total_area'));
    const pricePerSquare = Number(value('price_per_square'));
    const initial_payment = Number(value('initial_payment'));
    const totalPrice = pricePerSquare * total_area;
    const initialPaymentPercent = calculateInitialPaymentPercent(totalPrice, initial_payment);

    const propertyType = value('property_type');
    const isApartmentType = propertyType === 'apartment';
    const isCommerceType = propertyType === 'commerce';
    const isParkingType = propertyType === 'parking';

    // Общая часть для всех типов недвижимости
    const baseData = {
        isCashContract: value('contract') === 'cash',
        propertyType: propertyType,
        isApartmentType: isApartmentType,
        isCommerceType: isCommerceType,
        isParkingType: isParkingType,
        fullname: value('fullname'),
        shortname: getShortName(value('fullname')),
        birthdate: formatDate(value('birthdate')),
        phone_number: value('phone_number'),
        passport: value('passport'),
        passport_issue_date: formatDate(value('passport_issue_date')),
        passport_issued_by: value('passport_issued_by'),
        passport_division_code: value('passport_division_code'),
        registration_address: value('registration_address'),
        living_address: value('living_address'),
        installment_period: value('installment_period'),
        price_per_square: formatNumberWithSpaces(pricePerSquare),
        initial_payment: formatNumberWithSpaces(initial_payment),
        initial_payment_percent: initialPaymentPercent,
        price: formatNumberWithSpaces(totalPrice),

        current_date: formatDate(value('current_date'), true),
    };

    if (isParkingType) {
        return {
            ...baseData,
            parking_num: value('parking_num'),
        };
    }

    // Квартира и коммерция: одинаковый набор полей объекта,
    // отличие только в количестве комнат — у коммерции его нет
    const objectData = {
        ...baseData,
        building: value('building'),
        construction_number: value('construction_number'),
        floor: value('floor'),
        area: value('area'),
    };

    if (isApartmentType) {
        objectData.rooms = value('rooms');
    }

    return objectData;
}

// Шаблоны договоров: тип недвижимости -> тип договора
const CONTRACT_TEMPLATES = {
    apartment: {
        cash: 'template.docx',
        installment: 'template-installment.docx',
    },
    commerce: {
        cash: 'template-commerce.docx',
        installment: 'template-installment-commerce.docx',
    },
    parking: {
        cash: 'template-parking.docx',
        installment: 'template-parking-installment.docx',
    },
};

// Заполнение DOCX шаблона
async function fillDocxTemplate(data) {
    const paymentData = data.isCashContract ? {} : getPaymentDataForWord();

    const templates = CONTRACT_TEMPLATES[data.propertyType] || CONTRACT_TEMPLATES.apartment;
    const templateName = data.isCashContract ? templates.cash : templates.installment;

    const response = await fetch(templateName);
    const templateBuffer = await response.arrayBuffer();
    
    const zip = new PizZip(templateBuffer);

    // nullGetter: чтобы отсутствующий тег (например {rooms} в коммерции)
    // не подставлял в документ строку "undefined"
    const doc = new docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        nullGetter: () => '',
    });

    // Заполняем шаблон данными
    const documentData = {
        ...data,
        ...paymentData
    };
    doc.render(documentData);
    
    // Генерируем заполненный DOCX
    return doc.getZip().generate({type: 'arraybuffer'});
}

// Скачивание файла
function downloadFile(blob, fileName) {
    saveAs(blob, fileName);
}

function formatDate(dateString, isLong = false) {
    if (!dateString) return '';
    
    const date = new Date(dateString);
    
    if (isLong) {
        // Длинный формат: "22 Сентября 2025г."
        const day = date.getDate();
        const month = date.getMonth();
        const year = date.getFullYear();
        
        const monthNames = [
            'Января', 'Февраля', 'Марта', 'Апреля', 'Мая', 'Июня',
            'Июля', 'Августа', 'Сентября', 'Октября', 'Ноября', 'Декабря'
        ];
        
        return `${day} ${monthNames[month]} ${year}г.`;
    } else {
        // Короткий формат: "22.09.2025"
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        const year = String(date.getFullYear());
        
        return `${day}.${month}.${year}`;
    }
}

// Функция для форматирования валюты
function formatCurrency(amount) {
    return new Intl.NumberFormat('ru-RU').format(Math.round(amount));
}

function calculateInitialPaymentPercent(totalPrice, initialPayment) {
    if (totalPrice <= 0 || initialPayment <= 0) return 0;
    
    const percent = (initialPayment / totalPrice) * 100;
    return Math.ceil(percent);
}

async function convertDocxToPdf(docxBuffer) {
    return new Blob([docxBuffer], {type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
}

function formatNumberWithSpaces(number) {
    return new Intl.NumberFormat('ru-RU').format(number);
}

function getShortName(fullName) {
    if (!fullName) return '';
    
    const parts = fullName.trim().split(/\s+/);
    
    if (parts.length < 2) return fullName; 
    
    const lastName = parts[0];
    
    const initials = parts.slice(1)
        .map(name => name.charAt(0) + '.')
        .join(' ');
    
    return `${lastName} ${initials}`;
}

function displayPaymentPreview() {
    const contractType = document.getElementById('contract').value;
    const installmentPeriod = parseInt(document.getElementById('installment_period').value) || 0;
    
    if (contractType === 'installment' && installmentPeriod > 0) {
        const initialPayment = parseFloat(document.getElementById('initial_payment').value) || 0;
        const pricePerSquare = parseFloat(document.getElementById('price_per_square').value) || 0;
        const totalArea = parseFloat(document.getElementById('total_area').value) || 0;
        let currentDate = document.getElementById('current_date').value;
        const installmentCheckbox = document.getElementById('supportCheckbox').checked;

        if(installmentCheckbox) {
            currentDate = new Date(currentDate);
            currentDate.setMonth(currentDate.getMonth() + 1);
        }

        const payments = generatePaymentTable(initialPayment, pricePerSquare, totalArea, installmentPeriod, currentDate);

        // Показываем превью таблицы на странице
        const previewContainer = document.getElementById('payment-preview');
        if (previewContainer) {
            previewContainer.innerHTML = createPaymentTableHTML(payments);
        }
    }
}

function generatePaymentTable(initialPayment, pricePerSquare, totalArea, installmentPeriod, currentDate) {
    const totalPrice = pricePerSquare * totalArea;
    const installmentAmount = totalPrice - initialPayment;    
    const monthlyPayment = Math.round(installmentAmount / installmentPeriod);
    const payments = [];
    let remainingAmount = installmentAmount;
    const firstPaymentDate = new Date(currentDate);

    payments.push({
        number: 1,
        payment: formatCurrency(initialPayment),
        date: formatDate(firstPaymentDate),
        remaining: formatCurrency(installmentAmount)
    });

    for (let i = 0; i < installmentPeriod; i++) {
        const paymentDate = new Date(firstPaymentDate);
        paymentDate.setDate(1);
        paymentDate.setMonth(firstPaymentDate.getMonth() + (i + 1));

        const lastDayOfMonth = new Date(
            paymentDate.getFullYear(),
            paymentDate.getMonth() + 1,
            0
        ).getDate();

        if (firstPaymentDate.getDate() > lastDayOfMonth) {
            paymentDate.setDate(lastDayOfMonth);
        } else {
            paymentDate.setDate(firstPaymentDate.getDate());
        }

        // Форматируем дату в DD.MM.YYYY
        const formattedDate = formatDate(paymentDate);
        
        if (i === installmentPeriod - 1) {
            const lastPayment = remainingAmount;
            payments.push({
                number: i + 2,
                payment: formatCurrency(lastPayment),
                date: formattedDate,
                remaining: '0'
            });
        } else {
            remainingAmount -= monthlyPayment;
            
            payments.push({
                number: i + 2,
                payment: formatCurrency(monthlyPayment),
                date: formattedDate,
                remaining: formatCurrency(Math.max(0, remainingAmount))
            });
        }
    }
    
    return payments;
}

function createPaymentTableHTML(payments) {
    if (!payments || payments.length === 0) return '';
    
    let tableHTML = `
        <div style="margin-top: 40px; page-break-before: always;">
            <h3>График платежей</h3>
            <table border="1" cellpadding="8" cellspacing="0" style="width: 100%; border-collapse: collapse; font-size: 14px;">
                <thead>
                    <tr style="background-color: #f8f9fa;">
                        <th style="text-align: center; padding: 10px;">№ п/п</th>
                        <th style="text-align: center; padding: 10px;">Платеж (руб)</th>
                        <th style="text-align: center; padding: 10px;">Дата платежа</th>
                        <th style="text-align: center; padding: 10px;">Остаток (руб)</th>
                    </tr>
                </thead>
                <tbody>
    `;
    
    payments.forEach(payment => {
        tableHTML += `
            <tr>
                <td style="text-align: center; padding: 8px;">${payment.number}</td>
                <td style="text-align: right; padding: 8px;">${payment.payment}</td>
                <td style="text-align: center; padding: 8px;">${payment.date}</td>
                <td style="text-align: right; padding: 8px;">${payment.remaining}</td>
            </tr>
        `;
    });
    
    tableHTML += `
                </tbody>
            </table>
        </div>
    `;
    
    return tableHTML;
}

function getPaymentDataForWord() {
    const installmentPeriod = parseInt(document.getElementById('installment_period').value) || 0;

    const initialPayment = parseFloat(document.getElementById('initial_payment').value) || 0;
    const pricePerSquare = parseFloat(document.getElementById('price_per_square').value) || 0;
    const totalArea = parseFloat(document.getElementById('total_area').value) || 0;
    let currentDate = document.getElementById('current_date').value;
    const installmentCheckbox = document.getElementById('supportCheckbox').checked;

    if(installmentCheckbox) {
        currentDate = new Date(currentDate);
        currentDate.setMonth(currentDate.getMonth() + 1);
    }

    const payments = generatePaymentTable(initialPayment, pricePerSquare, totalArea, installmentPeriod, currentDate);
    
    const paymentTable = payments.map(payment => ({
        number: payment.number.toString(),
        payment: payment.payment,
        date: payment.date,
        remaining: payment.remaining
    }));
    
    return {
        paymentTable: paymentTable,
    };
}