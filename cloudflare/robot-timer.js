// robot-timer — a Cloudflare Worker that starts the Vic Air robots on time.
//
// GitHub's own schedule is best-effort and runs hours late, so this timer
// presses "Run workflow" on GitHub at fixed Melbourne times instead. It only
// does the starting: the robots, the figures and the SharePoint files stay
// where they are. Each start carries timer=true, so the robot follows its
// scheduled rules (fill only a missing slot; one email a day, weekdays).
// GitHub's schedule stays switched on underneath as the backup.
//
// Cloudflare setup:
//   - Cron trigger:  */5 * * * *   (every 5 minutes; the table below picks the times)
//   - Secret:        GITHUB_TOKEN  (fine-grained token, Actions: Read and write,
//                                   on the two repositories below only)
// Times are Melbourne local time, so daylight saving needs no change.

const OWNER = "VasDesign1";
const SALES = { repo: "bc-sales-tool", workflow: "snapshot.yml", ref: "main" };
const FORECAST = { repo: "forecast-tool", workflow: "snapshot.yml", ref: "master" };
const EMAIL = { repo: "bc-sales-tool", workflow: "daily-email.yml", ref: "main", weekdaysOnly: true };

const TIMETABLE = {
    // Morning snapshot (slot due 5:00), with two retries that do nothing if it is already there
    "05:00": [SALES], "05:10": [FORECAST],
    "05:30": [SALES], "05:40": [FORECAST],
    "06:00": [SALES], "06:10": [FORECAST],
    // Morning email: 6:30; the later two only act if the earlier one did not send
    "06:30": [EMAIL], "06:45": [EMAIL], "07:00": [EMAIL],
    // Midday snapshot (slot due 11:00)
    "11:00": [SALES], "11:10": [FORECAST],
    "11:30": [SALES], "11:40": [FORECAST],
    // Evening snapshot (slot due 17:00)
    "17:00": [SALES], "17:10": [FORECAST],
    "17:30": [SALES], "17:40": [FORECAST],
};

function melbourne(ms) {
    const parts = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", hour12: false, weekday: "short", hour: "2-digit", minute: "2-digit" }).formatToParts(new Date(ms));
    const g = t => parts.find(p => p.type === t).value;
    return { hhmm: String(parseInt(g("hour"), 10) % 24).padStart(2, "0") + ":" + g("minute"), weekday: g("weekday") };
}

async function start(job, token) {
    const r = await fetch("https://api.github.com/repos/" + OWNER + "/" + job.repo + "/actions/workflows/" + job.workflow + "/dispatches", {
        method: "POST",
        headers: { "Authorization": "Bearer " + token, "Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28", "User-Agent": "robot-timer", "Content-Type": "application/json" },
        body: JSON.stringify({ ref: job.ref, inputs: { timer: "true" } }),
    });
    if (r.status !== 204) throw new Error(job.repo + "/" + job.workflow + ": GitHub answered " + r.status + (r.status === 401 ? " (the access token has expired or is wrong)" : ""));
    console.log("Started " + job.repo + "/" + job.workflow);
}

export default {
    async scheduled(event, env) {
        const mel = melbourne(event.scheduledTime);
        const jobs = (TIMETABLE[mel.hhmm] || []).filter(j => !(j.weekdaysOnly && (mel.weekday === "Sat" || mel.weekday === "Sun")));
        if (!jobs.length) return;
        if (!env.GITHUB_TOKEN) throw new Error("GITHUB_TOKEN secret is not set");
        const results = await Promise.allSettled(jobs.map(j => start(j, env.GITHUB_TOKEN)));
        const failed = results.filter(r => r.status === "rejected");
        if (failed.length) throw new Error(failed.map(f => f.reason.message).join("; "));
    },
    // Nothing to see on the web address.
    async fetch() { return new Response("Not found", { status: 404 }); },
};
