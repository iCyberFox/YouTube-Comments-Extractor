    document.addEventListener('DOMContentLoaded', function() {
      const apiKey = import.meta.env.VITE_YOUTUBE_API_KEY;.
      let comments = [];
      const loadBtn = document.getElementById('loadBtn');
      const saveBtn = document.getElementById('saveBtn');
      const videoInput = document.getElementById('videoUrl');
      const errorMsg = document.getElementById('errorMsg');
      const tableBody = document.querySelector('#commentsTable tbody');
      
      loadBtn.addEventListener('click', loadComments);
      saveBtn.addEventListener('click', exportToExcel);
      
      function getVideoId(url) {
        if (!url) return null;
        
        const regExp = /^.*(youtu\.be\/|youtube\.com\/(shorts\/|embed\/|v\/|watch\?v=|watch\?.+&v=))([^#&?]*).*/;
        const match = url.match(regExp);
        
        return (match && match[3].length === 11) ? match[3] : null;
      }
      
      async function loadComments() {
        try {
          errorMsg.textContent = '';
          const videoId = getVideoId(videoInput.value.trim());
          
          if (!videoId) {
            showError('Будь ласка, введіть коректне посилання на YouTube');
            return;
          }
          
          tableBody.innerHTML = '<tr><td colspan="4" class="status">Завантаження коментарів...</td></tr>';
          comments = [];
          
          let nextPageToken = '';
          let counter = 1;
          
          do {
            const url = `https://www.googleapis.com/youtube/v3/commentThreads?` +
                       `part=snippet&videoId=${videoId}&key=${API_KEY}` +
                       `&maxResults=100${nextPageToken ? `&pageToken=${nextPageToken}` : ''}`;
            
            const response = await fetch(url);
            
            if (!response.ok) {
              const errorData = await response.json();
              throw new Error(errorData.error.message || 'Помилка API');
            }
            
            const data = await response.json();
            
            if (!data.items || data.items.length === 0) {
              if (comments.length === 0) {
                tableBody.innerHTML = '<tr><td colspan="4" class="status">Коментарі не знайдено</td></tr>';
              }
              return;
            }
            
            data.items.forEach(item => {
              const comment = item.snippet.topLevelComment.snippet;
              comments.push({
                number: counter++,
                name: comment.authorDisplayName,
                text: comment.textDisplay,
                date: new Date(comment.publishedAt).toLocaleDateString('uk-UA')
              });
            });
            
            nextPageToken = data.nextPageToken || '';
          } while (nextPageToken);
          
          renderComments();
          
        } catch (error) {
          console.error('Помилка:', error);
          showError('Помилка: ' + error.message);
          tableBody.innerHTML = '<tr><td colspan="4" class="status">Помилка завантаження</td></tr>';
        }
      }
      
      function renderComments() {
        tableBody.innerHTML = comments.map(comment => `
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
          showError('Немає даних для експорту');
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
            { width: 40 }, // Comment
            { width: 15 }   // Date
          ];

          // Отримуємо діапазон даних
          const range = XLSX.utils.decode_range(worksheet['!ref']);
          
          // Центруємо заголовки та вміст для Number і Date
          for (let C = range.s.c; C <= range.e.c; ++C) {
            for (let R = range.s.r; R <= range.e.r; ++R) {
              const cell = XLSX.utils.encode_cell({r: R, c: C});
              
              // Створюємо об'єкт стилю, якщо його ще немає
              if (!worksheet[cell].s) worksheet[cell].s = {};
              
              // Для всіх заголовків (перший рядок)
              if (R === range.s.r) {
                worksheet[cell].s.alignment = { 
                  horizontal: 'center',
                  vertical: 'center'
                };
              }
              // Для стовпців Number (0) і Date (3)
              else if (C === 0 || C === 3) {
                worksheet[cell].s.alignment = { 
                  horizontal: 'center',
                  vertical: 'center'
                };
              }
            }
          }

          const workbook = XLSX.utils.book_new();
          XLSX.utils.book_append_sheet(workbook, worksheet, 'YouTube Comments');
          XLSX.writeFile(workbook, `YouTube_Comments_${new Date().toISOString().slice(0,10)}.xlsx`);
          
        } catch (error) {
          console.error('Помилка експорту:', error);
          showError('Помилка при експорті до Excel');
        }
      }
      
      function showError(message) {
        errorMsg.textContent = message;
      }
    });