# Clear records recent files in MS Office for Mac

## English

I’m sure many people wanted to clear the list of previously opened files that is shown when launching office applications on a Mac.

<img width="866" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/ff08f599-de42-44e9-b548-e50392e60711">

Of course, it is possible to remove links to files one by one, but what if there are so many of them that deletion will take a lot of time?

<details>
   <summary>Spoiler How I reached the solution</summary>
   
I searched for information on this issue for a long time, and found a recommendation to delete the files:

```
com.microsoft.Excel.securebookmarks.plist
com.microsoft.Word.securebookmarks.plist
com.microsoft.PowerPoint.securebookmarks.plist
```

as described, let's say [here](https://answers.microsoft.com/en-us/msoffice/forum/all/how-to-clear-recent-files-list-in-one-go/abebd56a-91eb-453a-87ae-703e680cb325)
but it doesn't work on the latest versions of MS Office for Mac!

So where does Microsoft store the list of recently opened files these days?

While searching for information on the Internet, I came across a [recommendation](https://answers.microsoft.com/en-us/msoffice/forum/all/clear-and-disable-recent-files-in-office-365-on/063bcfcc-ee12-4a03-8124-8412c4204ca1) from a Microsoft employee which advised the user to delete the `MicrosoftRegistrationDB.reg` file, although it was said that all office application settings would be deleted along with the list of recent files... :facepalm:

Unfortunately, it was not possible to find it using the path indicated in the post, but I found the alias `MicrosoftRegistrationDB.reg`, which refers, as I think, to the file `MicrosoftRegistrationDB_xxxxxxxxxxxx.reg`, unique for each user (where x is a sequence of numbers), the second way to quickly find this file will probably be direct go to the folder along the path `~/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB`.

So this file is nothing more than a database in SQLite format.
Then it was a matter of time to understand how it stores records about the latest files and how they could be deleted.

As a result, I made a request to delete records relating to the last opened files:

`DELETE from "HKEY_CURRENT_USER_values" where node_id in (SELECT node_id FROM "HKEY_CURRENT_USER_values" WHERE name="path")`
</details>

To clear the entire list at once, do the following:
1. Open Finder, then in the “Go” menu select “Go to folder” and paste this path `~/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB`
2. We see one single file `MicrosoftRegistrationDB_xxxxxxxxxxxx.reg`
3. Open a terminal and type: `sqllite3` press space and drag the file from Finder into the terminal window and press Enter, SQLite will start and the database will open. You can make sure that the file with the database is open by typing `.tables`, you should see this output:
 
<img width="1204" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/e420625e-21e4-4d5d-9bd8-b09af46c6dc8">

4. You can check that the database contains your last opened files with the following request:

`select * from HKEY_CURRENT_USER_values where node_id in (SELECT node_id FROM HKEY_CURRENT_USER_values WHERE name='path');`[^1]

<img width="916" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/7d066ff1-0bcc-4a15-8374-5883300281e0">

5. You can actually clear this list with the following query:

`DELETE from "HKEY_CURRENT_USER_values" where node_id in (SELECT node_id FROM "HKEY_CURRENT_USER_values" WHERE name="path");`

<img width="928" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/fa15c7bf-0608-4b25-a9f0-88d21f9a5057">

6. Close SQLite with the `.quit` command.

<img width="265" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/90023452-2ff6-4854-a624-65d3f14d0953">

7. Launch Word, Excel, PowerPoint and enjoy the result!

PS You can modify the request to delete only the last opened documents of a specific application, to do this, you need to add a condition to the request for the "Value" field:

`Delete From HKEY_CURRENT_USER_values where node_id in (SELECT node_id FROM HKEY_CURRENT_USER_values WHERE value='Word');`

I know that you can wrap this whole algorithm in the form of a tasty script, but unfortunately I don’t know how to get the path to `MicrosoftRegistrationDB_xxxxxxxxxxxx.reg` from the alias `~/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB.reg`.

## Русский

Уверен многим хотелось очистить список ранее открытых файлов, который показывается при запуске офисных приложений на Mac.

<img width="866" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/ff08f599-de42-44e9-b548-e50392e60711">

Конечно есть возможность убирать ссылки на файлы по одному, но что если их так много, что на удаление уйдет много времени?

<details>
  <summary>Спойлер Как я дошел до решения</summary>
   
Я долго искал информацию по этому вопросу, нашел рекомендацию удалять файлы:

```
com.microsoft.Excel.securebookmarks.plist
com.microsoft.Word.securebookmarks.plist
com.microsoft.PowerPoint.securebookmarks.plist
```

как описано скажем вот [тут](https://answers.microsoft.com/en-us/msoffice/forum/all/how-to-clear-recent-files-list-in-one-go/abebd56a-91eb-453a-87ae-703e680cb325)
но это не работает на последних версиях MS Office for Mac!

Так где же Microsoft хранит нынче список последних открытых файлов?

В поисках информации в интернете я наткнулся на одну [рекомендацию](https://answers.microsoft.com/en-us/msoffice/forum/all/clear-and-disable-recent-files-in-office-365-on/063bcfcc-ee12-4a03-8124-8412c4204ca1) сорудника Microsoft, 
которая советовала пользователю удалить файл `MicrosoftRegistrationDB.reg`, правда было сказано что вместе со списком последних файлов удалятся все настройки офисных приложений... :facepalm:

К сожалению по указанному в посте пути обнаружить его не удалось, зато я нашел псевдоним `MicrosoftRegistrationDB.reg` ссылающийся, как я думаю, на уникальный у каждого пользователя файл `MicrosoftRegistrationDB_xxxxxxxxxxxx.reg` (где x это последовательность цифр), вторым способом быстро найти этот файл наверное будет прямой переход в папку по пути `~/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB`.
Так вот этот файл не что иное как база данных формата SQLite.

Дальше было делом времени понять каким образом в нем хранятся записи о послених файлах и каким образом их можно было бы удалить.

В итоге мною было составлен запрос на удаление записей касающихся последних открытых файлов:
`DELETE from "HKEY_CURRENT_USER_values" where node_id in (SELECT node_id FROM "HKEY_CURRENT_USER_values" WHERE name="path");`
</details>

Для очистки сразу всего списка необходимо сделать следующее:

1. Открываем Finder, далее в меню "Переход" выбираем пункт "Переход к папке" и вставляем этот путь ~/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB
2. Видим один единственный файл MicrosoftRegistrationDB_xxxxxxxxxxxx.reg
3. Открываем терминал и набираем: `sqllite3` нажимаем пробел и перетаскиваем файл из Finder в окно терминала и нажимаем Enter, запустится SQLite и откроется база. Убедится что файл с базой открыт можно набрав `.tables`, вы должны увидеть вот такой вывод:

<img width="1204" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/e420625e-21e4-4d5d-9bd8-b09af46c6dc8">

4. Проверить что в базе есть ваши последние открытые файлы можно следующим запросом:

`select * from HKEY_CURRENT_USER_values where node_id in (SELECT node_id FROM HKEY_CURRENT_USER_values WHERE name='path');`[^1]

<img width="916" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/7d066ff1-0bcc-4a15-8374-5883300281e0">

5. Собственно очистить этот список можно следующим запросом:

`DELETE from "HKEY_CURRENT_USER_values" where node_id in (SELECT node_id FROM "HKEY_CURRENT_USER_values" WHERE name="path");`

<img width="928" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/fa15c7bf-0608-4b25-a9f0-88d21f9a5057">

6. Закрываем SQLite командой `.quit`.

<img width="265" alt="image" src="https://github.com/idle4you/ClearRecordsRecentFilesMSOffice/assets/71770803/90023452-2ff6-4854-a624-65d3f14d0953">

7. Запускаем Word, Excel, PowerPoint и наслаждаемся результатом!

PS Можно модифицировать запрос чтобы удалить только последние открытые документы конкретного приложения, для этого необходимо добавить условие в запрос по полю "Value":

`Delete From HKEY_CURRENT_USER_values where node_id in (SELECT node_id FROM HKEY_CURRENT_USER_values WHERE value='Word');`

Знаю что можно весь этот алгоритм завернуть в виде вкусного скрипта, но к сожалению я не знаю как получить путь до MicrosoftRegistrationDB_xxxxxxxxxxxx.reg из псевдонима `~/Library/Group Containers/UBF8T346G9.Office/MicrosoftRegistrationDB.reg`.

[^1]: Semicolon at the end is required! Точка с запятой в конце обязательна!
