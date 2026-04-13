export function ConnectGoogleButton() {
    return (
        <a
            // href="/api/oauth/google/start"
            href="/api/api/oauth/google/start"
            style={{
                padding: 10,
                borderRadius: 10,
                color: "black",
                textAlign: "left",
                display: "block",
                textDecoration: "none"
            }}
        >
            Connect Google
        </a>
    );
}