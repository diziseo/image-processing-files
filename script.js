const { google } = require('googleapis');
const axios = require('axios');
const authenticate = require('./auth');
const fs = require('fs');
const { dialog, app, shell } = require('@electron/remote');
const { ipcRenderer } = require('electron');
const path = require('path');

const SHEET_ID = '1VikbryykNlac4PwR7b97n-1zyUIadcWC96v5cum8uRU';
const CORE_PATH = path.join(app.getPath('appData'), 'Ảnh by Dizi', 'core-files');
const CONFIG_PATH = path.join(CORE_PATH, 'config.json');
const INDICES_PATH = path.join(CORE_PATH, 'image_indices.json');

// Đọc config.json
const config = JSON.parse(fs.readFileSync(CONFIG_PATH));
const DRIVE_FOLDER_ID = config.driveFolderId;

// ... (Giữ nguyên phần còn lại của script.js)

let isEmailChecked = false;
let lastValidEmail = null;

function normalizeFileName(str) {
  return str
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, '-')
    .replace(/[^a-z0-9-]/g, '');
}

function loadImageIndices() {
  if (fs.existsSync(INDICES_PATH)) {
    const data = JSON.parse(fs.readFileSync(INDICES_PATH));
    return {
      backgroundIndex: data.backgroundIndex || 0,
      elementIndex: data.elementIndex || 0
    };
  }
  return { backgroundIndex: 0, elementIndex: 0 };
}

function saveImageIndices(backgroundIndex, elementIndex) {
  fs.writeFileSync(INDICES_PATH, JSON.stringify({ backgroundIndex, elementIndex }));
  console.log(`Đã lưu vị trí ảnh: backgroundIndex=${backgroundIndex}, elementIndex=${elementIndex}`);
}

async function uploadImageToCloudinary(fileBuffer, fileName, cloudinaryConfig) {
  const formData = new FormData();
  formData.append("file", new Blob([fileBuffer]), fileName);
  formData.append("upload_preset", config.cloudinary.uploadPreset);
  const response = await axios.post(
    `https://api.cloudinary.com/v1_1/${cloudinaryConfig.cloud_name}/image/upload`,
    formData,
    { validateStatus: () => true }
  );
  if (response.status >= 400) {
    console.log('Lỗi upload ảnh:', response.status, response.data);
    throw new Error(`Upload ảnh thất bại: ${response.statusText}`);
  }
  console.log('Upload ảnh thành công:', response.data.public_id);
  return response.data;
}

async function loadAdBanner() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!G2:G3',
    });
    const values = response.data.values || [];
    if (values.length >= 2) {
      const adImageUrl = values[0][0] || '';
      const adRedirectUrl = values[1][0] || '';
      const adBanner = document.getElementById('adBanner');
      if (adImageUrl) {
        const img = document.createElement('img');
        img.src = adImageUrl;
        img.onclick = () => adRedirectUrl && shell.openExternal(adRedirectUrl);
        adBanner.innerHTML = '';
        adBanner.appendChild(img);
      } else {
        adBanner.innerHTML = 'Không có ảnh quảng cáo';
      }
    } else {
      document.getElementById('adBanner').innerHTML = 'Không có dữ liệu quảng cáo';
    }
  } catch (error) {
    console.error('Lỗi tải quảng cáo từ G2/G3:', error);
    document.getElementById('adBanner').innerHTML = 'Lỗi tải quảng cáo';
  }
}

async function loadAdText() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!G1',
    });
    let adText = response.data.values ? response.data.values[0][0] : 'Đặt hàng liên hệ tele: @diziseo88';
    const adTextElement = document.getElementById('adText');
    adTextElement.classList.remove('marquee', 'color');
    let hasMarquee = adText.includes('[marquee]');
    let hasColor = adText.includes('[color]');
    adText = adText.replace('[marquee]', '').replace('[color]', '').trim();
    if (hasMarquee) adTextElement.classList.add('marquee');
    if (hasColor) adTextElement.classList.add('color');
    adTextElement.innerText = adText;
  } catch (error) {
    console.error('Lỗi tải quảng cáo:', error);
    document.getElementById('adText').innerText = 'Đặt hàng liên hệ tele: @diziseo88';
  }
}

async function loadServerAndElementData() {
  try {
    const auth = await authenticate();
    const sheets = google.sheets({ version: 'v4', auth });
    const backgroundResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!R2:S',
    });
    const backgroundData = (backgroundResponse.data.values || []).filter(row => row[0] && row[1]).map(row => ({ name: row[0], id: row[1] }));

    const serverResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!M2:P',
    });
    const serverData = (serverResponse.data.values || []).filter(row => row[0] && row[1] && row[2] && row[3]).map(row => ({ name: row[0], cloud_name: row[1], api_key: row[2], api_secret: row[3] }));

    const elementResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: 'Sheet1!J2:K',
    });
    const elementData = (elementResponse.data.values || []).filter(row => row[0] && row[1]).map(row => ({ name: row[0], id: row[1] }));

    const serverSelect = document.getElementById('serverSelect');
    serverSelect.innerHTML = '';
    serverData.forEach(server => {
      const option = document.createElement('option');
      option.value = server.name;
      option.text = server.name;
      serverSelect.appendChild(option);
    });

    const backgroundSelect = document.getElementById('backgroundSelect');
    backgroundSelect.innerHTML = '<option value="">Chọn nền</option>';
    backgroundData.forEach(background => {
      const option = document.createElement('option');
      option.value = background.name;
      option.text = background.name;
      backgroundSelect.appendChild(option);
    });

    const elementSelect = document.getElementById('elementSelect');
    elementSelect.innerHTML = '';
    elementData.forEach(element => {
      const option = document.createElement('option');
      option.value = element.name;
      option.text = element.name;
      elementSelect.appendChild(option);
    });

    return { serverData, backgroundData, elementData };
  } catch (error) {
    console.error('Lỗi tải dữ liệu:', error);
    return { serverData: [], backgroundData: [], elementData: [] };
  }
}

loadAdBanner();
loadAdText();
loadServerAndElementData().then(({ serverData, backgroundData, elementData }) => {
  const backgroundSelect = document.getElementById('backgroundSelect');
  const elementSelect = document.getElementById('elementSelect');
  const skipContentCheckbox = document.getElementById('skipContentCheckbox');
  const skipElementCheckbox = document.getElementById('skipElementCheckbox');
  const customImageInput = document.getElementById('customImageInput');
  const uploadCustomImageBtn = document.getElementById('uploadCustomImageBtn');
  const customElementImageInput = document.getElementById('customElementImageInput');
  const uploadCustomElementImageBtn = document.getElementById('uploadCustomElementImageBtn');

  let customImagePath = null;
  let customElementImagePath = null;

  function updateDisableState() {
    const isCustomImageSelected = !!customImagePath;
    const isCustomElementSelected = !!customElementImagePath;
    backgroundSelect.disabled = isCustomImageSelected;
    elementSelect.disabled = isCustomElementSelected;
    skipContentCheckbox.disabled = false;
    skipElementCheckbox.disabled = false;
  }

  updateDisableState();
  backgroundSelect.addEventListener('change', updateDisableState);
  elementSelect.addEventListener('change', updateDisableState);
  skipContentCheckbox.addEventListener('change', updateDisableState);
  skipElementCheckbox.addEventListener('change', updateDisableState);

  uploadCustomImageBtn.addEventListener('click', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
      title: 'Chọn ảnh nền từ máy tính',
      filters: [{ name: 'Image Files', extensions: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'] }],
      properties: ['openFile']
    });
    if (!canceled) {
      customImagePath = filePaths[0];
      customImageInput.value = customImagePath;
      updateDisableState();
    }
  });

  uploadCustomElementImageBtn.addEventListener('click', async () => {
    const { canceled, filePaths } = await dialog.showOpenDialog({
      title: 'Chọn ảnh phần tử từ máy tính',
      filters: [{ name: 'Image Files', extensions: ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'] }],
      properties: ['openFile']
    });
    if (!canceled) {
      customElementImagePath = filePaths[0];
      customElementImageInput.value = customElementImagePath;
      updateDisableState();
    }
  });

  document.getElementById('runButton').addEventListener('click', async () => {
    const statusText = document.getElementById('statusText');
    const loadingBarContainer = document.getElementById('loadingBarContainer');
    const loadingBar = document.getElementById('loadingBar');
    const expiryDateText = document.getElementById('expiryDate');
    const skipElement = skipElementCheckbox.checked;
    const skipContent = skipContentCheckbox.checked;

    statusText.innerText = 'Đang kiểm tra email...';
    loadingBarContainer.style.display = 'block';
    loadingBar.style.width = '0%';

    try {
      const auth = await authenticate();
      const sheets = google.sheets({ version: 'v4', auth });
      const drive = google.drive({ version: 'v3', auth });

      const emailInput = document.getElementById('emailInput').value.trim();
      ipcRenderer.send('set-current-email', emailInput);
      const logoUrl = document.getElementById('logoUrl').value;
      const serverName = document.getElementById('serverSelect').value;
      const backgroundFolderName = backgroundSelect.value;
      const elementFolderName = elementSelect.value;
      let contentInput = document.getElementById('contentInput').value;

      if (!emailInput) throw new Error("Chưa nhập email!");
      if (!logoUrl) throw new Error("Chưa nhập URL logo!");

      let data = skipContent ? [''] : contentInput.split('\n').filter(line => line.trim() !== '');
      if (!data.length && !skipElement) throw new Error("Chưa nhập nội dung!");

      let emailsColumnB = [];
      let emailsColumnF = [];
      let isTrial = false;

      if (!isEmailChecked) {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID,
          range: 'Sheet1!B:F',
        });
        const rows = response.data.values || [];
        emailsColumnB = rows.map(row => row[0] ? row[0].trim() : '').filter(email => email);
        emailsColumnF = rows.map(row => row[4] ? row[4].trim() : '').filter(email => email);
        const isEmailValid = emailsColumnB.includes(emailInput) && !emailsColumnF.includes(emailInput);

        if (emailsColumnB.includes(emailInput)) {
          const emailRowIndex = emailsColumnB.indexOf(emailInput);
          expiryDateText.innerText = `Ngày hết hạn: ${rows[emailRowIndex][2] || 'Không xác định'}`;
        } else {
          expiryDateText.innerText = 'Ngày hết hạn: Không tìm thấy email';
        }

        if (!emailsColumnB.includes(emailInput)) {
          isTrial = true;
          statusText.innerText = 'Mời bạn dùng thử 1 lần';
          data = [data[0] || ''];
        } else if (!isEmailValid) {
          throw new Error('Email đã được sử dụng!');
        } else {
          statusText.innerText = 'Đang chạy...';
          const emailRowIndex = emailsColumnB.indexOf(emailInput);
          rows[emailRowIndex][4] = emailInput;
          await sheets.spreadsheets.values.update({
            spreadsheetId: SHEET_ID,
            range: 'Sheet1!B1:F' + rows.length,
            valueInputOption: 'RAW',
            resource: { values: rows }
          });
        }
        isEmailChecked = true;
        lastValidEmail = emailInput;
      } else if (emailInput !== lastValidEmail) {
        throw new Error('Email không khớp với email đã kiểm tra!');
      } else {
        statusText.innerText = 'Đang chạy...';
      }

      const selectedServer = serverData.find(server => server.name === serverName);
      if (!selectedServer) throw new Error("Không tìm thấy server!");
      const cloudinaryConfig = selectedServer;

      let backgroundFolderId = null;
      let elementFolderId = null;
      let customBackgroundRes = null;
      let customElementRes = null;

      if (customImagePath) {
        customBackgroundRes = await uploadImageToCloudinary(fs.readFileSync(customImagePath), path.basename(customImagePath), cloudinaryConfig);
      } else if (backgroundFolderName) {
        const selectedBackground = backgroundData.find(bg => bg.name === backgroundFolderName);
        if (!selectedBackground) throw new Error("Không tìm thấy nền!");
        backgroundFolderId = selectedBackground.id;
      } else {
        throw new Error("Vui lòng chọn nền!");
      }

      if (!skipElement) {
        if (customElementImagePath) {
          customElementRes = await uploadImageToCloudinary(fs.readFileSync(customElementImagePath), path.basename(customElementImagePath), cloudinaryConfig);
        } else if (elementFolderName) {
          const selectedElement = elementData.find(el => el.name === elementFolderName);
          if (!selectedElement) throw new Error("Không tìm thấy phần tử!");
          elementFolderId = selectedElement.id;
        } else {
          throw new Error("Vui lòng chọn phần tử!");
        }
      }

      const supportedImageTypes = ["image/jpeg", "image/png", "image/gif", "image/bmp", "image/webp"];
      let backgroundArray = [];
      let elementArray = [];

      if (backgroundFolderId) {
        const backgroundFiles = await drive.files.list({
          q: `'${backgroundFolderId}' in parents`,
          fields: 'files(id, name, mimeType)',
        });
        backgroundArray = backgroundFiles.data.files.filter(file => supportedImageTypes.includes(file.mimeType));
        if (!backgroundArray.length) throw new Error("Không tìm thấy ảnh nền trong thư mục!");
      } else if (customBackgroundRes) {
        backgroundArray = [{ id: customBackgroundRes.public_id, name: path.basename(customImagePath), mimeType: customBackgroundRes.format }];
      }

      if (!skipElement && elementFolderId) {
        const elementFiles = await drive.files.list({
          q: `'${elementFolderId}' in parents`,
          fields: 'files(id, name, mimeType)',
        });
        elementArray = elementFiles.data.files.filter(file => supportedImageTypes.includes(file.mimeType));
        if (!elementArray.length) throw new Error("Không tìm thấy ảnh phần tử!");
      } else if (!skipElement && customElementRes) {
        elementArray = [{ id: customElementRes.public_id, name: path.basename(customElementImagePath), mimeType: customElementRes.format }];
      }

      const { backgroundIndex: startBackgroundIndex, elementIndex: startElementIndex } = loadImageIndices();
      let currentBackgroundIndex = startBackgroundIndex;
      let currentElementIndex = startElementIndex;

      const { canceled, filePath } = await dialog.showSaveDialog({
        title: 'Chọn vị trí lưu ảnh',
        defaultPath: 'output_image.webp',
        filters: [{ name: 'WebP Images', extensions: ['webp'] }]
      });
      if (canceled) throw new Error("Đã hủy chọn thư mục lưu!");
      const outputDir = path.dirname(filePath);

      for (let i = 0; i < data.length; i++) {
        const content = data[i];
        const normalizedContent = normalizeFileName(content || 'no-content');
        const backgroundFile = backgroundArray[currentBackgroundIndex % backgroundArray.length];
        console.log(`Xử lý ảnh ${i + 1}/${data.length} - Nền: ${backgroundFile.name}`);

        const progress = ((i + 1) / data.length) * 100;
        loadingBar.style.width = `${progress}%`;

        let backgroundRes = customBackgroundRes;
        if (backgroundFolderId) {
          const backgroundBlob = await drive.files.get({ fileId: backgroundFile.id, alt: 'media' }, { responseType: 'arraybuffer' });
          backgroundRes = await uploadImageToCloudinary(backgroundBlob.data, backgroundFile.name, cloudinaryConfig);
        }

        let elementRes = customElementRes;
        if (!skipElement && elementFolderId) {
          const elementFile = elementArray[currentElementIndex % elementArray.length];
          console.log(`Phần tử: ${elementFile.name}`);
          const elementBlob = await drive.files.get({ fileId: elementFile.id, alt: 'media' }, { responseType: 'arraybuffer' });
          elementRes = await uploadImageToCloudinary(elementBlob.data, elementFile.name, cloudinaryConfig);
        }

        const logoRes = await axios.get(logoUrl, { responseType: 'arraybuffer', validateStatus: () => true });
        if (logoRes.status >= 400) throw new Error(`Tải logo thất bại: ${logoRes.statusText}`);
        const logoUploadRes = await uploadImageToCloudinary(logoRes.data, "logo.png", cloudinaryConfig);

        const elementWidth = Math.floor(backgroundRes.width * 0.9);
        const elementHeight = Math.floor(backgroundRes.height * 0.9);

        let transformUrl = `https://res.cloudinary.com/${cloudinaryConfig.cloud_name}/image/upload/q_50,f_webp`;
        transformUrl += `/l_${logoUploadRes.public_id},g_north_west,x_10,y_10,w_120`;
        if (!skipElement && elementRes) transformUrl += `/l_${elementRes.public_id},w_${elementWidth},h_${elementHeight},c_fit,g_center`;
        if (!skipContent) transformUrl += `/l_text:Roboto_28_bold:${encodeURIComponent(content.toUpperCase())},co_rgb:FFFFFF,g_south,x_0,y_20,b_rgb:000000`;
        transformUrl += `/${backgroundRes.public_id}.webp`;

        console.log(`URL ảnh: ${transformUrl}`);
        const finalImageRes = await axios.get(transformUrl, { responseType: 'arraybuffer', validateStatus: () => true });
        if (finalImageRes.status >= 400) throw new Error(`Tạo ảnh thất bại: ${finalImageRes.statusText}`);

        const outputPath = `${outputDir}/${normalizedContent}-${i}.webp`;
        fs.writeFileSync(outputPath, Buffer.from(finalImageRes.data));
        console.log(`Đã lưu ảnh: ${outputPath}`);

        currentBackgroundIndex++;
        if (!skipElement) currentElementIndex++;
      }

      saveImageIndices(currentBackgroundIndex % backgroundArray.length, skipElement ? startElementIndex : currentElementIndex % elementArray.length);

      loadingBarContainer.style.display = 'none';
      statusText.innerText = 'Mời thượng đế kiểm tra ảnh tại thư mục đã chọn';
      if (isTrial && !emailsColumnB.includes(emailInput)) app.quit();
    } catch (error) {
      console.error('Lỗi:', error);
      loadingBarContainer.style.display = 'none';
      statusText.innerText = 'Lỗi: ' + error.message;
    }
  });
});