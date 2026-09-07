# Remove Paywalls

### 1) Delete paywall elements in the page

### 2) Add scrolling functionality

JS code:

In developer tools (F12 in chrome), go to console and run this snippet:

    document.documentElement.style.setProperty('overflow', 'auto', 'important');
    document.body.style.setProperty('overflow', 'auto', 'important');
    document.body.style.setProperty('position', 'static', 'important');
    document.body.style.setProperty('height', 'auto', 'important');
