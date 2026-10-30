export default function handler(req, res) {
    res.status(200).json({
        scriptUrl: process.env.APPS_SCRIPT_URL
    });
}