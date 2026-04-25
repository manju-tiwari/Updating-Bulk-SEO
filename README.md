# Multi-Site WordPress SEO Meta Updater
## Works with: Yoast SEO | Rank Math | All in One SEO

---

## STEP 1 — Install Python Dependencies

```bash
pip install requests openpyxl
```

---

## STEP 2 — Generate WordPress App Passwords

For EACH website:
1. Log in to WordPress Admin
2. Go to: Users → Your Profile
3. Scroll to "Application Passwords" section
4. Enter name (e.g., "SEO Updater") → Click "Add New"
5. Copy the generated password (save it immediately!)

---

## STEP 3 — Configure seo_meta_updater.py

Open the script and edit the SITES section:

```python
SITES = {
    "site1": {
        "url": "https://yoursite1.com",
        "username": "admin",
        "app_password": "AbCd EfGh IjKl MnOp QrSt UvWx",
        "seo_plugin": "yoast"        # yoast | rankmath | aioseo
    },
    "site2": {
        "url": "https://yoursite2.com",
        "username": "admin",
        "app_password": "AbCd EfGh IjKl MnOp QrSt UvWx",
        "seo_plugin": "rankmath"
    },
    "site3": {
        "url": "https://yoursite3.com",
        "username": "admin",
        "app_password": "AbCd EfGh IjKl MnOp QrSt UvWx",
        "seo_plugin": "aioseo"
    },
}
```

---

## STEP 4 — Prepare Your CSV / Excel File

File name: seo_updates.csv  (or .xlsx)

### Required Columns:
| Column           | Description                              | Example             |
|------------------|------------------------------------------|---------------------|
| site_key         | Must match key in SITES config           | site1               |
| post_id          | WordPress post/page ID                   | 101                 |
| post_type        | post / page / product                    | post                |
| meta_title       | New SEO title                            | Best Laptops 2026   |
| meta_description | New meta description (150-160 chars)     | Discover the top... |

### How to find Post ID:
- Go to WordPress Admin → Posts/Pages
- Hover over any post → URL shows ?post=ID

---

## STEP 5 — Run the Script

```bash
python seo_meta_updater.py
```

Output example:
```
[INFO] Loaded 8 rows from seo_updates.csv
[INFO] [1/8] Processing: site=site1 | post_id=101
[INFO]   ✓ Updated successfully (Plugin: yoast)
[INFO] [2/8] Processing: site=site1 | post_id=102
[INFO]   ✓ Updated successfully (Plugin: yoast)
...
[INFO] SUMMARY
[INFO]   Total Rows   : 8
[INFO]   ✓ Successful : 8
[INFO]   ✗ Failed     : 0
[INFO]   ⚠ Skipped    : 0
```

---

## FIELD MAPPINGS (for reference)

| Plugin       | Meta Title Field          | Meta Desc Field           |
|--------------|---------------------------|---------------------------|
| Yoast SEO    | _yoast_wpseo_title        | _yoast_wpseo_metadesc     |
| Rank Math    | rank_math_title           | rank_math_description     |
| All in One   | _aioseo_title             | _aioseo_description       |

---

## TROUBLESHOOTING

| Error                    | Fix                                                  |
|--------------------------|------------------------------------------------------|
| 401 Unauthorized         | Check username & app_password in config              |
| 404 Not Found            | Verify post_id exists on that site                  |
| Cannot connect           | Check site URL (include https://)                   |
| Meta not updating        | Enable REST API meta exposure (see below)           |

### Enable Meta Exposure in REST API
Add this to your theme's functions.php or via Code Snippets plugin:

```php
// For Yoast
add_filter('is_protected_meta', function($protected, $meta_key) {
    $yoast_keys = ['_yoast_wpseo_title', '_yoast_wpseo_metadesc'];
    if (in_array($meta_key, $yoast_keys)) return false;
    return $protected;
}, 10, 2);

register_post_meta('post', '_yoast_wpseo_title', [
    'show_in_rest' => true, 'single' => true, 'type' => 'string',
    'auth_callback' => function() { return current_user_can('edit_posts'); }
]);
register_post_meta('post', '_yoast_wpseo_metadesc', [
    'show_in_rest' => true, 'single' => true, 'type' => 'string',
    'auth_callback' => function() { return current_user_can('edit_posts'); }
]);
```

---

## AUTOMATE ON SCHEDULE (Optional)

### Windows Task Scheduler:
- Action: python C:\path\to\seo_meta_updater.py
- Trigger: Daily / Weekly as needed

### Linux/Mac Cron:
```bash
# Run every Monday at 9 AM
0 9 * * 1 /usr/bin/python3 /path/to/seo_meta_updater.py
```
