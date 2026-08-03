/**
 * Перевіряє та створює всі відсутні каталоги (OU) по вказаному шляху.
 */
function archiveEnsureOrgUnit_(ouPath) {
    if (!ouPath || ouPath === "/") return;
    
    let s = String(ouPath).trim();
    if (!s.startsWith("/")) s = "/" + s;
    while (s.length > 1 && s.endsWith("/")) s = s.slice(0, -1);
    
    let existingOUs = [];
    let pageToken;
    do {
        let res = archiveCallWithRetry_(() => AdminDirectory.Orgunits.list("my_customer", { type: "all", pageToken: pageToken }));
        if (res.organizationUnits) {
            existingOUs = existingOUs.concat(res.organizationUnits);
        }
        pageToken = res.nextPageToken;
    } while (pageToken);
    
    if (existingOUs.some(ou => ou.orgUnitPath === s)) {
        return; 
    }
    
    const parts = s.split("/").filter(p => p.trim().length > 0);
    let parentPath = "";
    
    for (let i = 0; i < parts.length; i++) {
        const name = parts[i];
        const currentPath = parentPath === "" ? "/" + name : parentPath + "/" + name;
        
        let found = existingOUs.find(ou => ou.orgUnitPath === currentPath);
        
        if (!found) {
            let parentData = parentPath === "" ? { orgUnitPath: "/" } : existingOUs.find(ou => ou.orgUnitPath === parentPath);
            
            if (!parentData) {
                throw new Error("Батьківська OU не знайдена: " + parentPath);
            }
            
            let inserted = archiveCallWithRetry_(() => AdminDirectory.Orgunits.insert({
                name: name,
                parentOrgUnitPath: parentData.orgUnitPath
            }, "my_customer"));
            
            existingOUs.push(inserted);
        }
        
        parentPath = currentPath;
    }
}

/**
 * Допоміжна функція для безпечного виклику API з повторними спробами (Retry Pattern).
 */
function archiveCallWithRetry_(fn, maxRetries = 3) {
    let attempt = 0;
    while (true) {
        try {
            return fn();
        } catch (e) {
            const msg = (e && e.message) ? e.message : String(e);
            const transient = /Rate Limit|Quota|quota|Too many|Service invoked too many times|429|500|503|Backend Error/i.test(msg);
            
            if (!transient || attempt >= maxRetries) throw e;
            
            Utilities.sleep(1000 * Math.pow(2, attempt));
            attempt++;
        }
    }
}
