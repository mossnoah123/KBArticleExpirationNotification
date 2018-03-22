# KB Article Expiration Notification Script

## Summary
The purpose of this script is to automate the task of manually tracking when KB Articles should be reviewed and pushed for reactivation.

The [Teaching and Technology Center](https://www.uwplatt.edu/ttc) (TTC) at UW-Platteville keeps track of which KB Articles they are reponsible for in a Google Sheet. Using this sheet, the TTC was having to manually set reminders of when KB Articles were nearing their expiration date. However, with the power of Google's scripting language, this tedious task is now successfully being automated.

Below is a sample Google Sheet that this script interacts with. It shows the sheet that the TTC uses to keep track of all TTC-owned KB Articles.

![Google Sheet displaying all of the TTC-owned KB Articles](Assets/TTC%20KB%20Article%20Google%20Sheet%20Example.png "This is a sample Google Sheet; the TTC's KB Article Spreadsheet")

## Background

### UW KnowledgeBase
The University of Wisconsin System created and hosts a [KnowledgeBase](https://kb.wisc.edu/) (KB) platform for "easily creating, displaying, sharing, and managing web-based knowledge documents".<sup>**1**</sup> The schools in the University of Wisconsin System have the option to adopt a KB of their own that they can have attached to their school's domain. For example, [UW-Platteville has a KnowledgeBase](https://kb.uwplatt.edu).

### The Problem
In practice, when KB articles are created, they are alive for 1 year before they expire. When an article expires, it is up to the owner of the article to review it for outdated information; if they decide the information of the article is still relevant, then the article can be reactivated for another year. **The problem** with this is there does not exist any functionality for KB articles to automatically notify their owners when they are about to expire. It is ultimately up to the article owner and users with internal access to the KB to keep track of when articles expire, or are nearing expiration, so that the review process may begin. However, when there are hundreds of articles to manage, this task can be daunting and can go untouched for a significant amount of time. 

The downtime of an expired KB article with still-relevant and helpful information negatively impacts users who are searching for a particular article and can't find a solution to their problem. These users do not realize that what they are searching for may exist in an expired state within the internal KB and they simply cannot see this article for that reason. It is important that this information stays fresh, up-to-date and readily available for users.

### The Project
This is a script that I wrote in the Spring of 2018 for the [Teaching and Technology Center](https://www.uwplatt.edu/ttc) at UW-Platteville. During the creation of this script, I had barely touched JavaScript before, so this project was a learning experience. This project took me roughly a day and a half to create during a chilly, early Spring weekend.

My goal with this project was to finally get hands-on with JavaScript, something I had been wanting to do for awhile. I also wanted to learn more about the structure of a JavaScript file and its naming/documentation conventions. On a side note, it was fun approaching this project with an OO-dominant background and using what I know to structure the script well so that it could be easily maintained.

## Future Plans for this Project
* Eliminate the need for a Google Sheet altogether and use the [KnowledgeBase API](https://kb.wisc.edu/kbGuide/page.php?id=71945) to determine expiring articles.


## References
<sup>**1**</sup> - [About the KnowledgeBase](https://kb.wisc.edu/page.php?id=3)
