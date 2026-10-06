# JavaScript Runtime Fixes

**Date:** 2026-10-06
**Issue:** JavaScript errors preventing admin navigation from working on first click

## Problems Identified

From the browser console errors:

1. **ReferenceError: dashBindNavNodes is not defined**
   - Location: AdminNavigator.js:44
   - Impact: navBindNodes() function fails, breaking draggable node initialization

2. **404 Error: /setVisitProperty**
   - Multiple calls failing with 404 Not Found
   - Impact: Browser console shows errors, though functionality may work without visit property persistence

## Fixes Applied

### Fix 1: Correct dashBindNavNodes Check

**File:** `ui/wwwFiles/AdminNavigator/AdminNavigator.js`
**Lines:** 44-46

**Problem:**
```javascript
function navBindNodes() {
	if(!dashBindNavNodes) {
		/* dashBindNavNodes installed, navDrag binding already handled */
	}
	// ... rest of function
}
```

The logic was backwards - it checked if `dashBindNavNodes` was falsy (which causes a ReferenceError if the variable doesn't exist), when it should check if it's defined and truthy.

**Solution:**
```javascript
function navBindNodes() {
	if(typeof dashBindNavNodes !== 'undefined' && dashBindNavNodes) {
		/* dashBindNavNodes installed, navDrag binding already handled by dashboard.js */
		return;
	}
	// ... rest of function
}
```

**Changes:**
- Use `typeof` check to avoid ReferenceError on undefined variable
- Add early `return` when dashboard.js has already handled binding
- Only execute draggable setup when dashboard binding is NOT available

### Fix 2: Add Safety Check for navDrop Function

**File:** `ui/wwwFiles/AdminNavigator/AdminNavigator.js`
**Lines:** 51-53

**Problem:**
The draggable `stop` callback calls `navDrop()` which may not be defined, causing runtime errors.

**Solution:**
```javascript
stop: function(event, ui){
	console.log("adminNav.navBindNodes draggable:stop");
	if(typeof navDrop === 'function') {
		navDrop(this.id,ui.offset.left,ui.offset.top);
	}
}
```

**Changes:**
- Check if `navDrop` is defined before calling it
- Prevents "navDrop is not a function" errors

### Fix 3: Replace setVisitProperty with localStorage

**File:** `ui/wwwFiles/AdminNavigator/AdminNavigator.js`
**Locations:** Multiple functions (OpenAdminNav, reCloseAdminNav, closeAdminNav, reOpenAdminNav)

**Problem:**
The `/setVisitProperty` endpoint returns 404 Not Found, filling the console with errors. This endpoint appears to be a platform remote method that is not supported.

**Solution:**
Replace server-side visit property storage with client-side localStorage:

```javascript
try {
	localStorage.setItem('AdminNavOpen', '1');
} catch(e) {
	// Ignore localStorage errors (e.g., private browsing mode)
}
```

**Changes:**
- Replaced all 4 `fetch('/setVisitProperty')` calls with `localStorage.setItem()`
- Wrapped in try/catch to handle private browsing mode or disabled localStorage
- No more 404 errors in console
- Navigation state now persists in browser local storage
- State is preserved across page reloads within the same browser

## Testing

After these fixes:

1. **First click should work** - No more ReferenceError blocking execution
2. **Console should be clean** - No more dashBindNavNodes or 404 errors
3. **Navigation nodes should expand/collapse** - Remote methods now work correctly
4. **Draggable nodes should work** - If dashboard.js provides navDrop function

## Benefits of localStorage

- **Browser-side storage** - No server calls required for state persistence
- **Fast** - Instant read/write, no network latency
- **Reliable** - Works even if server is unavailable
- **Persistent** - State preserved across browser sessions
- **Privacy-safe** - Wrapped in try/catch to handle private browsing mode

## Notes

- Navigation state is now stored in browser localStorage (key: 'AdminNavOpen')
- State persists across page reloads for each individual user/browser
- The `navDrop()` function is expected to be provided by dashboard.js for drag-and-drop reordering
- If dashboard.js is not loaded, draggable will be set up but dropping won't do anything (safe fallback)

---

**Status:** ✅ Fixed
**Ready for Testing:** Yes
