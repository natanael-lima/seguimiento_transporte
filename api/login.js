export default function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Método no permitido' });
    }

    const { password } = req.body;
    const serverPassword = process.env.ADMIN_PASSWORD;

    if (password === serverPassword) {
        return res.status(200).json({ success: true, message: 'Acceso autorizado' });
    } else {
        return res.status(401).json({ success: false, message: 'Contraseña incorrecta' });
    }
}