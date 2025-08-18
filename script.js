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
        // Отримуємо API ключ з змінної середовища Netlify
        const response = await fetch('/.netlify/functions/getApiKey');
        if (response.ok) {
          const data = await response.json();
          API_KEY = data.apiKey;
          apiStatus.textContent = 'API ключ успішно завантажено';
          loadBtn.disabled = false;
        } else {
          throw new Error('Не вдалося отримати API ключ');
        }
      } catch (error) {
        console.error('Помилка отримання API ключа:', error);
        apiStatus.textContent = 'Помилка завантаження API ключа';
        apiStatus.classList.add('error');
        return;
      }
      
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
          
          let nextPageToken = '';
          let index = 1;
          let hasComments = false;
          
          do {
            const response = await fetch(
              `https://www.googleapis.com/youtube/v3/commentThreads?` +
              `part=snippet&videoId=${videoId}&key=${API_KEY}` +
              `&maxResults=100${nextPageToken ? `&pageToken=${nextPageToken}` : ''}`
            );
            
            if (!response.ok) {
              const errorData = await response.json();
              let errorMessage = 'Помилка при отриманні коментарів';
              if (errorData.error && errorData.error.message) {
                errorMessage += ': ' + errorData.error.message;
                if (errorData.error.errors && errorData.error.errors[0].reason === 'commentsDisabled') {
                  errorMessage = 'Коментарі вимкнені для цього відео';
                }
              }
              throw new Error(errorMessage);
            }
            
            const data = await response.json();
            
            if (!data.items || data.items.length === 0) {
              if (!hasComments) {
                showMessage("Коментарі не знайдено");
              }
              return;
            }
            
            hasComments = true;
            data.items.forEach(item => {
              const comment = item.snippet.topLevelComment.snippet;
              comments.push({
                number: index++,
                name: comment.authorDisplayName,
                text: comment.textDisplay,
                date: new Date(comment.publishedAt).toLocaleDateString('uk-UA')
              });
            });
            
            nextPageToken = data.nextPageToken || '';
          } while (nextPageToken);
          
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