using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace GitHubUploaderApp
{
    public class UploaderForm : Form
    {
        TextBox txtToken;
        TextBox txtRepo;
        CheckBox chkPrivate;

        CheckBox chkCreateRelease;
        TextBox txtTag;
        TextBox txtReleaseTitle;
        TextBox txtReleaseNotes;
        CheckBox chkAttachExe;
        CheckBox chkAttachZip;

        Button btnUpload;
        Button btnPaste;
        ProgressBar progress;
        TextBox txtLog;
        string _projectDir;

        const string DefaultReleaseNotes =
@"### 🚀 Новое в версии 1.1.0:

#### ⌨ Свобода раскладки клавиатуры и исправление ввода:
- **Свободный ввод без принудительной раскладки:** Программа больше не перехватывает язык Windows и не сбрасывает раскладку на русскую при вводе данных, переходе между полями или открытии окон. Вы можете комфортно вводить текст на любой раскладке.
- **Корректные адреса с дробями, точками и запятыми:** Исправлена ошибка, из-за которой не сохранялись специальные символы в номерах домов и квартир. Теперь адреса вида `1/7`, `3.4`, `8,3`, а также литерные номера (`дом №17А`) сохраняются в точности так, как введено.
- **Удобный ввод слэша на русской раскладке:** При наборе адреса на русской раскладке клавиша `\` (обратный слэш) автоматически подставляет `/` (прямой слэш). Переключать язык ради одного символа больше не нужно!
- **Интеллектуальная транслитерация текста:** Сохранено исправление случайного набора ФИО и улиц латинскими буквами («dmitriev» ➔ «ДМИТРИЕВ»), при этом алгоритм стал полностью безопасным для цифр, дробей и знаков препинания.

#### 🎨 Новый современный и понятный интерфейс:
- **Карточная компоновка полей:** Все реквизиты сгруппированы в аккуратные карточки:
  - 👤 *Должники и представитель* (номер дела, выбор префикса, список должников, заявитель и взыскатель).
  - 🏠 *Адрес объекта недвижимости* (улица с авто-припиской, дом в формате «дом №...», квартира/комната с поддержкой дробей).
  - 💰 *Суммы к взысканию* (основной долг, пени, госпошлина и автоматический расчет итоговой суммы).
- **Чистая шапка окна генератора:** Убраны громоздкие технические пути к файлам. Выбор образца оформлен компактной кнопкой `⚙ Шаблон...`.
- **Быстрое автозаполнение тестами:** На панель действий добавлена заметная кнопка `🎲 Тестовые данные` (горячие клавиши `Ctrl+Shift+T` и `F12`) для мгновенного заполнения формы реальными демонстрационными реквизитами для проверки генерации.
- **Крупная кнопка сохранения:** Акцентная кнопка `✔ Создать заявление (Enter)` обеспечивает максимальную скорость работы оператора.
- **Быстрый доступ к папкам:** В главном окне добавлены кнопки быстрого открытия папок «Новые заявления», «Исходные заявления» и «Исправленные заявления».

#### 🛠 Надежность и переносимость:
- 100% совместимость с эталонным образцом Word `094729862.docx`.
- Автономная работа на любом ПК с Windows без установки дополнительных библиотек (.NET Framework 4.8 / 4.0).
- Быстрая компиляция в 1 клик через `СБОРКА ПРОГРАММЫ.bat`.";

        public UploaderForm()
        {
            _projectDir = AppDomain.CurrentDomain.BaseDirectory.TrimEnd('\\', '/');

            Text = "Загрузка проекта и публикация релиза на GitHub";
            Size = new Size(760, 740);
            MinimumSize = new Size(700, 600);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9.5f);
            BackColor = Color.FromArgb(248, 249, 250);

            try
            {
                string iconPath = Path.Combine(_projectDir, "app.ico");
                if (File.Exists(iconPath)) Icon = new Icon(iconPath);
            }
            catch { }

            BuildUi();
        }

        void BuildUi()
        {
            var pnlTop = new Panel
            {
                Dock = DockStyle.Top,
                Height = 54,
                BackColor = Color.FromArgb(36, 41, 47),
                Padding = new Padding(16, 10, 16, 10)
            };

            var lblHeader = new Label
            {
                Text = "Публикация проекта и создание релиза на GitHub",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlTop.Controls.Add(lblHeader);

            var pnlScroll = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(16)
            };

            int y = 10;

            // --- Блок 1: Настройки подключения ---
            var grpAuth = new GroupBox
            {
                Text = "  1. Настройки репозитория GitHub  ",
                Location = new Point(16, y),
                Size = new Size(700, 140),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41)
            };

            int gy = 24;
            grpAuth.Controls.Add(new Label
            {
                Text = "GitHub Personal Access Token (с правом 'repo'):",
                Location = new Point(14, gy),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            });
            gy += 22;

            txtToken = new TextBox
            {
                Location = new Point(14, gy),
                Size = new Size(540, 26),
                Font = new Font("Consolas", 10f),
                UseSystemPasswordChar = true
            };
            grpAuth.Controls.Add(txtToken);

            btnPaste = new Button
            {
                Text = "Вставить",
                Location = new Point(564, gy - 1),
                Size = new Size(118, 28),
                Font = new Font("Segoe UI", 9f)
            };
            btnPaste.Click += (s, e) =>
            {
                if (Clipboard.ContainsText())
                {
                    txtToken.Text = Clipboard.GetText().Trim();
                    AppendLog("Токен вставлен из буфера обмена.");
                }
            };
            grpAuth.Controls.Add(btnPaste);
            gy += 32;

            grpAuth.Controls.Add(new Label
            {
                Text = "Имя репозитория:",
                Location = new Point(14, gy + 3),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            });

            txtRepo = new TextBox
            {
                Text = "pomoshnik-zayavleniya",
                Location = new Point(140, gy),
                Size = new Size(240, 26),
                Font = new Font("Segoe UI", 9.5f)
            };
            grpAuth.Controls.Add(txtRepo);

            chkPrivate = new CheckBox
            {
                Text = "Приватный репозиторий",
                Location = new Point(410, gy + 2),
                AutoSize = true,
                Checked = false,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            };
            grpAuth.Controls.Add(chkPrivate);

            pnlScroll.Controls.Add(grpAuth);
            y += 150;

            // --- Блок 2: Настройки релиза ---
            var grpRelease = new GroupBox
            {
                Text = "  2. Публикация релиза (GitHub Release для скачивания)  ",
                Location = new Point(16, y),
                Size = new Size(700, 285),
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 37, 41)
            };

            int ry = 24;
            chkCreateRelease = new CheckBox
            {
                Text = "Создать официальный GitHub Release (со страницей скачивания и списком изменений)",
                Location = new Point(14, ry),
                AutoSize = true,
                Checked = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.DarkSlateBlue
            };
            chkCreateRelease.CheckedChanged += (s, e) =>
            {
                txtTag.Enabled = chkCreateRelease.Checked;
                txtReleaseTitle.Enabled = chkCreateRelease.Checked;
                txtReleaseNotes.Enabled = chkCreateRelease.Checked;
                chkAttachExe.Enabled = chkCreateRelease.Checked;
                chkAttachZip.Enabled = chkCreateRelease.Checked;
            };
            grpRelease.Controls.Add(chkCreateRelease);
            ry += 28;

            grpRelease.Controls.Add(new Label
            {
                Text = "Тег версии:",
                Location = new Point(14, ry + 3),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            });

            txtTag = new TextBox
            {
                Text = "v1.1.0",
                Location = new Point(95, ry),
                Size = new Size(95, 26),
                Font = new Font("Segoe UI", 9.5f)
            };
            grpRelease.Controls.Add(txtTag);

            grpRelease.Controls.Add(new Label
            {
                Text = "Название релиза:",
                Location = new Point(205, ry + 3),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            });

            txtReleaseTitle = new TextBox
            {
                Text = "Помощник по заявлениям v1.1.0 — Исправление раскладки и новый интерфейс",
                Location = new Point(325, ry),
                Size = new Size(355, 26),
                Font = new Font("Segoe UI", 9.5f)
            };
            grpRelease.Controls.Add(txtReleaseTitle);
            ry += 32;

            var lblNotes = new Label
            {
                Text = "Описание релиза / Что нового (Markdown):",
                Location = new Point(14, ry + 4),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            };
            grpRelease.Controls.Add(lblNotes);

            var btnPresetV11 = new Button
            {
                Text = "✨ Данные v1.1.0",
                Location = new Point(365, ry),
                Size = new Size(130, 26),
                Font = new Font("Segoe UI", 8.5f),
                BackColor = Color.FromArgb(235, 243, 255),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnPresetV11.FlatAppearance.BorderColor = Color.FromArgb(180, 205, 240);
            btnPresetV11.Click += (s, e) =>
            {
                txtTag.Text = "v1.1.0";
                txtReleaseTitle.Text = "Помощник по заявлениям v1.1.0 — Исправление раскладки и новый интерфейс";
                txtReleaseNotes.Text = DefaultReleaseNotes;
                AppendLog("Загружены стандартные данные релиза v1.1.0.");
            };
            grpRelease.Controls.Add(btnPresetV11);

            var btnLoadFile = new Button
            {
                Text = "📂 Из release_notes.txt",
                Location = new Point(505, ry),
                Size = new Size(175, 26),
                Font = new Font("Segoe UI", 8.5f),
                BackColor = Color.FromArgb(245, 245, 245),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnLoadFile.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btnLoadFile.Click += (s, e) =>
            {
                string defFile = Path.Combine(_projectDir, "release_notes.txt");
                if (File.Exists(defFile) && (ModifierKeys & Keys.Shift) != Keys.Shift)
                {
                    LoadReleaseInfoFromFile(defFile, true);
                }
                else
                {
                    using (var ofd = new OpenFileDialog())
                    {
                        ofd.InitialDirectory = _projectDir;
                        ofd.Filter = "Текстовые файлы (*.txt;*.md)|*.txt;*.md|Все файлы (*.*)|*.*";
                        ofd.FileName = "release_notes.txt";
                        ofd.Title = "Выберите файл с описанием релиза";
                        if (ofd.ShowDialog(this) == DialogResult.OK)
                        {
                            LoadReleaseInfoFromFile(ofd.FileName, true);
                        }
                    }
                }
            };
            grpRelease.Controls.Add(btnLoadFile);
            ry += 30;

            txtReleaseNotes = new TextBox
            {
                Location = new Point(14, ry),
                Size = new Size(668, 120),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Text = DefaultReleaseNotes,
                Font = new Font("Segoe UI", 9f)
            };
            grpRelease.Controls.Add(txtReleaseNotes);
            ry += 128;

            chkAttachExe = new CheckBox
            {
                Text = "Прикрепить к релизу готовый файл 'Помощник по заявлениям.exe'",
                Location = new Point(14, ry),
                AutoSize = true,
                Checked = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            };
            grpRelease.Controls.Add(chkAttachExe);

            chkAttachZip = new CheckBox
            {
                Text = "Прикрепить полный ZIP-архив релиза",
                Location = new Point(445, ry),
                AutoSize = true,
                Checked = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular)
            };
            grpRelease.Controls.Add(chkAttachZip);

            pnlScroll.Controls.Add(grpRelease);
            y += 295;

            // --- Кнопка запуска ---
            btnUpload = new Button
            {
                Text = "🚀  Загрузить проект и опубликовать Релиз на GitHub",
                Location = new Point(16, y),
                Size = new Size(700, 42),
                BackColor = Color.FromArgb(46, 164, 79),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat
            };
            btnUpload.FlatAppearance.BorderSize = 0;
            btnUpload.Click += async (s, e) => await StartUploadAsync();
            pnlScroll.Controls.Add(btnUpload);
            y += 50;

            progress = new ProgressBar
            {
                Location = new Point(16, y),
                Size = new Size(700, 8),
                Visible = false,
                Style = ProgressBarStyle.Marquee
            };
            pnlScroll.Controls.Add(progress);
            y += 14;

            txtLog = new TextBox
            {
                Location = new Point(16, y),
                Size = new Size(700, 130),
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                BackColor = Color.White,
                Font = new Font("Consolas", 9f),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };
            pnlScroll.Controls.Add(txtLog);

            Controls.Add(pnlScroll);
            Controls.Add(pnlTop);

            // Автоматическая загрузка описания из release_notes.txt, если файл существует
            string autoReleaseFile = Path.Combine(_projectDir, "release_notes.txt");
            if (File.Exists(autoReleaseFile))
            {
                LoadReleaseInfoFromFile(autoReleaseFile, false);
            }
        }

        void LoadReleaseInfoFromFile(string filePath, bool notify)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    if (notify)
                        MessageBox.Show("Файл не найден:\n" + filePath, "Загрузка данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
                string tag = null;
                string title = null;
                var bodyLines = new List<string>();
                bool inBody = false;

                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (!inBody)
                    {
                        if (line.StartsWith("TAG:", StringComparison.OrdinalIgnoreCase))
                        {
                            tag = line.Substring(4).Trim();
                            continue;
                        }
                        if (line.StartsWith("TITLE:", StringComparison.OrdinalIgnoreCase))
                        {
                            title = line.Substring(6).Trim();
                            continue;
                        }
                        if (line.Trim() == "---")
                        {
                            inBody = true;
                            continue;
                        }
                    }
                    bodyLines.Add(line);
                }

                if (!string.IsNullOrEmpty(tag)) txtTag.Text = tag;
                if (!string.IsNullOrEmpty(title)) txtReleaseTitle.Text = title;

                string body = string.Join("\r\n", bodyLines.ToArray()).Trim();
                if (!string.IsNullOrEmpty(body))
                {
                    txtReleaseNotes.Text = body;
                }

                AppendLog("Загружены данные релиза из файла: " + Path.GetFileName(filePath));
                if (notify)
                {
                    MessageBox.Show(
                        string.Format("Данные релиза успешно загружены!\n\nФайл: {0}\nТег: {1}\nНазвание: {2}",
                            Path.GetFileName(filePath), txtTag.Text, txtReleaseTitle.Text),
                        "Данные релиза загружены",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                AppendLog("Ошибка чтения файла релиза: " + ex.Message);
                if (notify)
                    MessageBox.Show("Ошибка при чтении файла:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void AppendLog(string msg)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AppendLog), msg);
                return;
            }
            txtLog.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] " + msg + "\r\n");
        }

        async Task StartUploadAsync()
        {
            string token = txtToken.Text.Trim();
            if (string.IsNullOrEmpty(token))
            {
                MessageBox.Show("Введите ваш Personal Access Token от GitHub.", "Требуется токен",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtToken.Focus();
                return;
            }

            string repoName = txtRepo.Text.Trim();
            if (string.IsNullOrEmpty(repoName))
            {
                MessageBox.Show("Укажите имя репозитория.", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtRepo.Focus();
                return;
            }

            string tag = txtTag.Text.Trim();
            if (chkCreateRelease.Checked && string.IsNullOrEmpty(tag))
            {
                MessageBox.Show("Укажите тег версии (например: v1.0.0).", "Внимание",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTag.Focus();
                return;
            }

            btnUpload.Enabled = false;
            btnPaste.Enabled = false;
            progress.Visible = true;

            bool isPrivate = chkPrivate.Checked;
            bool createRelease = chkCreateRelease.Checked;
            string releaseTitle = txtReleaseTitle.Text.Trim();
            string releaseNotes = txtReleaseNotes.Text;
            bool attachExe = chkAttachExe.Checked;
            bool attachZip = chkAttachZip.Checked;

            try
            {
                string resultUrl = await Task.Run(() => UploadWorker(token, repoName, isPrivate, createRelease, tag, releaseTitle, releaseNotes, attachExe, attachZip));
                AppendLog("✔ УСПЕХ! Всё готово: " + resultUrl);

                var res = MessageBox.Show(
                    "Проект и релиз успешно опубликованы на GitHub!\n\nСсылка:\n" + resultUrl + "\n\nОткрыть страницу релиза в браузере?",
                    "Успешно опубликовано", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (res == DialogResult.Yes)
                {
                    try { System.Diagnostics.Process.Start(resultUrl); } catch { }
                }
            }
            catch (Exception ex)
            {
                AppendLog("ОШИБКА: " + ex.Message);
                MessageBox.Show("Ошибка: " + ex.Message, "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnUpload.Enabled = true;
                btnPaste.Enabled = true;
                progress.Visible = false;
            }
        }

        string UploadWorker(string token, string repoName, bool isPrivate,
            bool createRelease, string releaseTag, string releaseTitle, string releaseNotes,
            bool attachExe, bool attachZip)
        {
            var js = new JavaScriptSerializer();

            AppendLog("[1/6] Авторизация на GitHub...");
            string userJson = GitHubApi("GET", "https://api.github.com/user", token, null);
            var userObj = js.Deserialize<Dictionary<string, object>>(userJson);
            string owner = (string)userObj["login"];
            AppendLog(string.Format("Успешно авторизован как: {0}", owner));

            AppendLog(string.Format("[2/6] Проверка репозитория {0}/{1}...", owner, repoName));
            string repoUrl = "";
            try
            {
                string repoJson = GitHubApi("GET", string.Format("https://api.github.com/repos/{0}/{1}", owner, repoName), token, null);
                var repoObj = js.Deserialize<Dictionary<string, object>>(repoJson);
                repoUrl = (string)repoObj["html_url"];
                AppendLog("Репозиторий найден: " + repoUrl);
            }
            catch
            {
                AppendLog("Создаём новый репозиторий: " + repoName);
                var createObj = new Dictionary<string, object>
                {
                    { "name", repoName },
                    { "private", isPrivate },
                    { "description", "Помощник по заявлениям (ООО «ЛУКОЙЛ-Волгоградэнерго»)" },
                    { "auto_init", true }
                };
                string createdJson = GitHubApi("POST", "https://api.github.com/user/repos", token, js.Serialize(createObj));
                var createdObj = js.Deserialize<Dictionary<string, object>>(createdJson);
                repoUrl = (string)createdObj["html_url"];
                AppendLog("Репозиторий успешно создан: " + repoUrl);
                System.Threading.Thread.Sleep(2000);
            }

            AppendLog("[3/6] Отправка исходных файлов проекта...");
            var filesToUpload = new List<string>
            {
                "App.cs",
                "app.ico",
                "build.ps1",
                "СБОРКА ПРОГРАММЫ.bat",
                "улицы.txt",
                "README.md",
                "release_notes.txt",
                ".gitignore",
                "GitHubUploader.cs",
                "Шаблоны\\094729862.docx",
                "Новые заявления\\.gitkeep",
                "Исходные заявления\\.gitkeep",
                "Исправленные заявления\\.gitkeep"
            };

            // Получаем ветку main или master
            string parentSha = null;
            string branchName = "main";
            try
            {
                string refJson = GitHubApi("GET", string.Format("https://api.github.com/repos/{0}/{1}/git/ref/heads/main", owner, repoName), token, null);
                var refObj = js.Deserialize<Dictionary<string, object>>(refJson);
                var objDict = (Dictionary<string, object>)refObj["object"];
                parentSha = (string)objDict["sha"];
            }
            catch
            {
                try
                {
                    string refJson = GitHubApi("GET", string.Format("https://api.github.com/repos/{0}/{1}/git/ref/heads/master", owner, repoName), token, null);
                    var refObj = js.Deserialize<Dictionary<string, object>>(refJson);
                    var objDict = (Dictionary<string, object>)refObj["object"];
                    parentSha = (string)objDict["sha"];
                    branchName = "master";
                }
                catch { }
            }

            var treeItems = new List<Dictionary<string, object>>();

            foreach (var relPath in filesToUpload)
            {
                string fullPath = Path.Combine(_projectDir, relPath);
                if (!File.Exists(fullPath)) continue;

                AppendLog("  Файл: " + relPath);
                byte[] bytes = File.ReadAllBytes(fullPath);
                string b64 = Convert.ToBase64String(bytes);

                var blobReq = new Dictionary<string, object>
                {
                    { "content", b64 },
                    { "encoding", "base64" }
                };

                string blobJson = GitHubApi("POST", string.Format("https://api.github.com/repos/{0}/{1}/git/blobs", owner, repoName), token, js.Serialize(blobReq));
                var blobObj = js.Deserialize<Dictionary<string, object>>(blobJson);
                string blobSha = (string)blobObj["sha"];

                treeItems.Add(new Dictionary<string, object>
                {
                    { "path", relPath.Replace('\\', '/') },
                    { "mode", "100644" },
                    { "type", "blob" },
                    { "sha", blobSha }
                });
            }

            AppendLog("[4/6] Фиксация коммита...");
            var treeReq = new Dictionary<string, object> { { "tree", treeItems } };
            string treeJson = GitHubApi("POST", string.Format("https://api.github.com/repos/{0}/{1}/git/trees", owner, repoName), token, js.Serialize(treeReq));
            var treeObj = js.Deserialize<Dictionary<string, object>>(treeJson);
            string treeSha = (string)treeObj["sha"];

            var commitReq = new Dictionary<string, object>
            {
                { "message", "Релиз " + (createRelease ? releaseTitle : "Помощник по заявлениям") },
                { "tree", treeSha },
                { "parents", parentSha != null ? new object[] { parentSha } : new object[0] }
            };

            string commitJson = GitHubApi("POST", string.Format("https://api.github.com/repos/{0}/{1}/git/commits", owner, repoName), token, js.Serialize(commitReq));
            var commitObj = js.Deserialize<Dictionary<string, object>>(commitJson);
            string commitSha = (string)commitObj["sha"];

            // Обновление ветки
            if (parentSha != null)
            {
                var updateRefReq = new Dictionary<string, object> { { "sha", commitSha }, { "force", true } };
                GitHubApi("PATCH", string.Format("https://api.github.com/repos/{0}/{1}/git/refs/heads/{2}", owner, repoName, branchName), token, js.Serialize(updateRefReq));
            }
            else
            {
                var newRefReq = new Dictionary<string, object> { { "ref", "refs/heads/main" }, { "sha", commitSha } };
                GitHubApi("POST", string.Format("https://api.github.com/repos/{0}/{1}/git/refs", owner, repoName), token, js.Serialize(newRefReq));
            }

            AppendLog(string.Format("[5/6] Ветка {0} успешно обновлена.", branchName));

            // --- Блок создания релиза ---
            if (createRelease)
            {
                AppendLog(string.Format("[6/6] Публикация официального релиза {0}...", releaseTag));

                int releaseId = 0;
                string releaseHtmlUrl = "";

                // Проверяем, существует ли уже релиз с таким тегом
                try
                {
                    string existingJson = GitHubApi("GET", string.Format("https://api.github.com/repos/{0}/{1}/releases/tags/{2}", owner, repoName, releaseTag), token, null);
                    var existingObj = js.Deserialize<Dictionary<string, object>>(existingJson);
                    releaseId = Convert.ToInt32(existingObj["id"]);
                    releaseHtmlUrl = (string)existingObj["html_url"];
                    AppendLog("Обновление существующего релиза: " + releaseHtmlUrl);

                    var editBody = new Dictionary<string, object>
                    {
                        { "name", releaseTitle },
                        { "body", releaseNotes },
                        { "draft", false },
                        { "prerelease", false }
                    };
                    GitHubApi("PATCH", string.Format("https://api.github.com/repos/{0}/{1}/releases/{2}", owner, repoName, releaseId), token, js.Serialize(editBody));
                }
                catch
                {
                    var releaseBody = new Dictionary<string, object>
                    {
                        { "tag_name", releaseTag },
                        { "target_commitish", branchName },
                        { "name", releaseTitle },
                        { "body", releaseNotes },
                        { "draft", false },
                        { "prerelease", false }
                    };

                    string releaseJson = GitHubApi("POST", string.Format("https://api.github.com/repos/{0}/{1}/releases", owner, repoName), token, js.Serialize(releaseBody));
                    var releaseObj = js.Deserialize<Dictionary<string, object>>(releaseJson);
                    releaseId = Convert.ToInt32(releaseObj["id"]);
                    releaseHtmlUrl = (string)releaseObj["html_url"];
                    AppendLog("Создан новый релиз: " + releaseHtmlUrl);
                }

                // Удаление старых ассетов релиза с такими же именами (или с искажёнными _._.exe), чтобы избежать ошибки 422
                try
                {
                    string assetsJson = GitHubApi("GET", string.Format("https://api.github.com/repos/{0}/{1}/releases/{2}/assets", owner, repoName, releaseId), token, null);
                    var assetsList = js.Deserialize<object[]>(assetsJson);
                    foreach (Dictionary<string, object> a in assetsList)
                    {
                        string aname = (string)a["name"];
                        if (aname == "Pomoshnik_po_zayavleniyam.exe" || aname == "_._.exe" || aname.StartsWith("_._") || aname.StartsWith("Pomoshnik_po_zayavleniyam_"))
                        {
                            int aid = Convert.ToInt32(a["id"]);
                            try { GitHubApi("DELETE", string.Format("https://api.github.com/repos/{0}/{1}/releases/assets/{2}", owner, repoName, aid), token, null); } catch { }
                        }
                    }
                }
                catch { }

                // Прикрепление файла Помощник по заявлениям.exe (используем безопасное имя Pomoshnik_po_zayavleniyam.exe)
                string exePath = Path.Combine(_projectDir, "Помощник по заявлениям.exe");
                if (attachExe && File.Exists(exePath))
                {
                    AppendLog("  Прикрепление к релизу: Pomoshnik_po_zayavleniyam.exe...");
                    byte[] exeBytes = File.ReadAllBytes(exePath);
                    string uploadUrl = string.Format("https://uploads.github.com/repos/{0}/{1}/releases/{2}/assets?name={3}",
                        owner, repoName, releaseId, "Pomoshnik_po_zayavleniyam.exe");
                    try { UploadBinaryAsset(uploadUrl, token, exeBytes); }
                    catch (Exception ex) { AppendLog("  Примечание (asset exe): " + ex.Message); }
                }

                // Прикрепление полного ZIP архива
                if (attachZip)
                {
                    string zipName = string.Format("Pomoshnik_po_zayavleniyam_{0}.zip", releaseTag);
                    AppendLog("  Создание и прикрепление архива: " + zipName + "...");
                    byte[] zipBytes = CreateReleaseZip();
                    string uploadUrl = string.Format("https://uploads.github.com/repos/{0}/{1}/releases/{2}/assets?name={3}",
                        owner, repoName, releaseId, zipName);
                    try { UploadBinaryAsset(uploadUrl, token, zipBytes); }
                    catch (Exception ex) { AppendLog("  Примечание (asset zip): " + ex.Message); }
                }

                return releaseHtmlUrl;
            }

            return repoUrl;
        }

        byte[] CreateReleaseZip()
        {
            using (var ms = new MemoryStream())
            {
                using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                {
                    var files = new[]
                    {
                        "Помощник по заявлениям.exe",
                        "README.md",
                        "release_notes.txt",
                        "улицы.txt",
                        "Шаблоны\\094729862.docx",
                        "СБОРКА ПРОГРАММЫ.bat",
                        "build.ps1",
                        "App.cs",
                        "app.ico"
                    };

                    foreach (var rel in files)
                    {
                        string fp = Path.Combine(_projectDir, rel);
                        if (!File.Exists(fp)) continue;

                        var entry = archive.CreateEntry(rel, CompressionLevel.Optimal);
                        using (var es = entry.Open())
                        using (var fs = File.OpenRead(fp))
                        {
                            fs.CopyTo(es);
                        }
                    }
                }
                return ms.ToArray();
            }
        }

        static string GitHubApi(string method, string url, string token, string jsonBody)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;

            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = method;
            req.UserAgent = "Pomoshnik-GitHub-Uploader";
            req.Accept = "application/vnd.github.v3+json";
            req.Headers.Add("Authorization", "token " + token);

            if (jsonBody != null)
            {
                req.ContentType = "application/json; charset=utf-8";
                byte[] b = Encoding.UTF8.GetBytes(jsonBody);
                req.ContentLength = b.Length;
                using (var s = req.GetRequestStream()) s.Write(b, 0, b.Length);
            }

            try
            {
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var reader = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    using (var reader = new StreamReader(wex.Response.GetResponseStream(), Encoding.UTF8))
                    {
                        string err = reader.ReadToEnd();
                        throw new Exception(string.Format("GitHub API ({0}): {1}", ((HttpWebResponse)wex.Response).StatusCode, err));
                    }
                }
                throw;
            }
        }

        static void UploadBinaryAsset(string uploadUrl, string token, byte[] data)
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls;
            var req = (HttpWebRequest)WebRequest.Create(uploadUrl);
            req.Method = "POST";
            req.UserAgent = "Pomoshnik-GitHub-Uploader";
            req.Accept = "application/vnd.github.v3+json";
            req.Headers.Add("Authorization", "token " + token);
            req.ContentType = "application/octet-stream";
            req.ContentLength = data.Length;

            using (var s = req.GetRequestStream())
            {
                s.Write(data, 0, data.Length);
            }

            try
            {
                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    // OK
                }
            }
            catch (WebException wex)
            {
                if (wex.Response != null)
                {
                    using (var reader = new StreamReader(wex.Response.GetResponseStream(), Encoding.UTF8))
                    {
                        string err = reader.ReadToEnd();
                        throw new Exception("Asset upload: " + err);
                    }
                }
                throw;
            }
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new UploaderForm());
        }
    }
}
