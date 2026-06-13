    document.addEventListener('DOMContentLoaded', async function() {
      const loadBtn = document.getElementById('loadBtn');
      const saveBtn = document.getElementById('saveBtn');
      const videoInput = document.getElementById('videoUrl');
      const errorMsg = document.getElementById('errorMsg');
      const tableBody = document.querySelector('#commentsTable tbody');
      const apiStatus = document.getElementById('apiStatus');
      
      let API_KEY = '';
      let comments = [];
      
      try {
        const response = await fetch('/.netlify/functions/getApiKey');
        if (response.ok) {
          const data = await response.json();
          API_KEY = data.apiKey;
        }
      } catch (error) {
        console.warn('API key check skipped:', error);
      }

      apiStatus.textContent = 'Готово до завантаження';
      apiStatus.classList.remove('error');
      loadBtn.disabled = false;
      
      loadBtn.addEventListener('click', loadComments);
      saveBtn.addEventListener('click', exportToExcel);
      
      function getVideoId(url) {
        if (!url) return null;
        
        try {
          // Для скорочених URL (youtu.be/VIDEO_ID)
          if (url.includes('youtu.be/')) {
            return url.split('youtu.be/')[1].split(/[?&#]/)[0];
          }
          
          // Для YouTube Shorts
          if (url.includes('youtube.com/shorts/')) {
            return url.split('youtube.com/shorts/')[1].split(/[?&#]/)[0];
          }
          
          // Для звичайних YouTube відео
          const regExp = /^.*(youtube\.com\/watch\?v=|youtube\.com\/embed\/|youtube\.com\/v\/|youtu\.be\/)([^#&?]*).*/;
          const match = url.match(regExp);
          
          return (match && match[2].length === 11) ? match[2] : null;
        } catch (e) {
          return null;
        }
      }
      
      async function loadComments() {
        try {
          errorMsg.textContent = '';
          const url = videoInput.value.trim();
          if (!url) {
            showError("Будь ласка, введіть посилання на відео");
            return;
          }

          const videoId = getVideoId(url);
          if (!videoId || videoId.length !== 11) {
            showError("Невірний формат посилання. Введіть коректне посилання на YouTube відео або Shorts");
            return;
          }

          showMessage("Завантаження коментарів...");
          comments = [];
          saveBtn.disabled = true;

          const commentsResponse = await fetch(`/.netlify/functions/getComments?videoId=${encodeURIComponent(videoId)}`);
          if (!commentsResponse.ok) {
            const errorData = await commentsResponse.json().catch(() => ({}));
            throw new Error(errorData.error || 'Помилка при отриманні коментарів');
          }

          const data = await commentsResponse.json();

          if (!data.comments || data.comments.length === 0) {
            showMessage("Коментарі не знайдено");
            return;
          }

          comments = data.comments.map((comment, index) => ({
            number: index + 1,
            name: comment.name,
            text: comment.text,
            date: comment.date
          }));

          renderTable();
          saveBtn.disabled = false;
        } catch (error) {
          console.error("Помилка:", error);
          showError(error.message);
          showMessage("Помилка завантаження");
          saveBtn.disabled = true;
        }
      }

      function showMessage(msg) {
        document.querySelector("#commentsTable tbody").innerHTML = `
          <tr>
            <td colspan="4" class="status">${msg}</td>
          </tr>
        `;
      }

      function renderTable() {
        const tbody = document.querySelector("#commentsTable tbody");
        if (comments.length === 0) {
          showMessage("Коментарі не знайдені");
          return;
        }
        
        tbody.innerHTML = comments.map(comment => `
          <tr>
            <td>${comment.number}</td>
            <td>${comment.name}</td>
            <td>${comment.text}</td>
            <td>${comment.date}</td>
          </tr>
        `).join('');
      }

      function exportToExcel() {
        if (comments.length === 0) {
          showError("Немає даних для експорту");
          return;
        }
        
        try {
          const excelData = comments.map(comment => ({
            'Number': comment.number,
            'Name': comment.name,
            'Comment': comment.text,
            'Date': comment.date
          }));

          const worksheet = XLSX.utils.json_to_sheet(excelData);
          
          // Налаштування ширини колонок
          worksheet['!cols'] = [
            { width: 8 },   // Number
            { width: 25 },  // Name
            { width: 40 },  // Comment
            { width: 15 }   // Date
          ];

          // Додаємо стилі для центрування
          const range = XLSX.utils.decode_range(worksheet['!ref']);
          
          // Центруємо заголовки
          for (let C = range.s.c; C <= range.e.c; ++C) {
            const headerCell = XLSX.utils.encode_cell({r: range.s.r, c: C});
            if (!worksheet[headerCell].s) worksheet[headerCell].s = {};
            worksheet[headerCell].s.alignment = { 
              horizontal: 'center',
              vertical: 'center'
            };
          }
          
          // Центруємо вміст для Number (0) і Date (3)
          for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            // Number column
            const numberCell = XLSX.utils.encode_cell({r: R, c: 0});
            if (!worksheet[numberCell].s) worksheet[numberCell].s = {};
            worksheet[numberCell].s.alignment = { 
              horizontal: 'center',
              vertical: 'center'
            };
            
            // Date column
            const dateCell = XLSX.utils.encode_cell({r: R, c: 3});
            if (!worksheet[dateCell].s) worksheet[dateCell].s = {};
            worksheet[dateCell].s.alignment = { 
              horizontal: 'center',
              vertical: 'center'
            };
          }

          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, 'YouTube Comments');
          XLSX.writeFile(workbook, `YouTube_Comments_${new Date().toISOString().slice(0,10)}.xlsx`);
          
        } catch (error) {
          console.error("Помилка експорту:", error);
          showError("Помилка при експорті до Excel");
        }
      }
      
      function showError(message) {
        errorMsg.textContent = message;
      }
    });