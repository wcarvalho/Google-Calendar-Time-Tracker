#### Time Management Script

<table>
  <tr>
    <td>
      <img class="regular materialboxed responsive-img" src="/img/ex_calendar_final_postweek.png">
    </td>
    <td>
      <img class="regular materialboxed responsive-img" src="/img/ex_spreadsheet_final_postweek.png">
    </td>
  </tr>
  <tr>
    <td>The calendar names are on the left, while the category names are used in the event titles.</td>
    <td>Time delegation for categories "570," "561," "research," and "other" was only tracked using calendars "Work," "Extra," and "Personal."</td>
  </tr>
</table>

<b>Menu:</b>

| Feature | Purpose |
| - | - |
| Fill Category Times | Fill in B3 onwards with the times for events in the categories above |
| Fill Calendar Times | Fill in C8 onwards with the time assigned to their corresponding calendars |
| Fill Calendar Row Colors | Color in the calendar rows according to the colors written in A8 onwards |

<br>
<b>Installation Instructions:</b>
<!-- ####Installation Instructions -->
<!-- A brief overview of how to use the script in my post, "A Prescription for Managing and Tracking Your Time," to track time delegation in Google Calendar using a Google spreadsheet -->

1. Download [this excel file](https://docs.google.com/spreadsheets/d/1ELRQ8M8bjhPlvydnJxGaMsTuwaKN6YKjQJLUZ3zmKFs/pub?output=xlsx), [upload it](https://docs.google.com/spreadsheets/u/0/) into Google Sheets, and add [this script](/files/time_management/time_tracker.js) to it.
  
    <i>Right-click script link, select "Save Link As...", and save as .txt file</i><br>
    <i>Once you upload the excel file, open "Tools -> Script Editor..."</i><br>
    <i>Paste the contents of the script inside and save.</i><br>
    <i>Close and re-open the tab and the menu should appear.</i>

2. Place the name of calendars you want to track vertically beginning at B8.
3. If you want a calendar to contribute to tracking a category's time assignment, place text in the field to the left of the calendar name. (If the text is a color (plain text or hexidecimal), you can use "Fill Calendar Row Colors" to color in the corresponding row.)
4. Place the name of categories you want to track horizontally beginning at B2.
5. An event will contribute to a category's time if the category is somewhere in the event's title.
6. Put the start date and end date across cells E7-H7 and E8-H8, respectively, in the format (date, month, year).

    <i>Script defaults to current time when start date is empty or invalid.</i><br>
    <i>Script defaults to Saturday 12:00am when end date is empty or invalid.</i>

7. Track your time using the Time Tracker menu which should appear as soon as you save the script.


**Good luck managing your time! Please leave any comments, suggestions, or questions below.**

#####For details on how I use this tool, check out [this post](/2016/01/08/TimeManagementPrescription.html).