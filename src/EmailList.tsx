import type {Email} from "./types.ts";

export interface EmailListProps {
    emails: Email[];
    onEmailClick: (email: Email) => void;
    selectedEmail: Email | null;
}

const buildPreviewBody = (body: string): string => {
    return body
        .split(/\r?\n/)
        .filter(line => !line.trimStart().startsWith('>'))
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
};

// In EmailList.tsx
export function EmailList({ emails, onEmailClick, selectedEmail }: EmailListProps) {
    return (
        <div style={{ background: "#f9f9f9" }}>
            {emails.map((email, idx) => {
                const previewBody = buildPreviewBody(email.body);
                const isSelected = selectedEmail &&
                    selectedEmail.from === email.from &&
                    selectedEmail.to === email.to &&
                    selectedEmail.subject === email.subject &&
                    selectedEmail.receivedAt === email.receivedAt;
                const isRead = email.read === true || email.unread === false;

                return (
                    <div
                        key={idx}
                        onClick={(e) => {
                            e.stopPropagation();
                            onEmailClick(email);
                        }}
                        style={{
                            padding: "15px 8px",
                            cursor: "pointer",
                            borderBottom: "1px solid #ddd",
                            background: isSelected ? "#d4e9ff" : "white",
                            opacity: isRead ? 0.6 : 1,
                            borderLeft: isSelected ? "3px solid #007bff" : "3px solid transparent"
                        }}
                    >
                        <div style={{
                            display: "flex",
                            justifyContent: "space-between",
                            alignItems: "center"
                        }}>
                            <strong style={{ fontWeight: isRead ? 400 : 600 }}>
                                {email.attachments && email.attachments.length > 0 && (
                                    <span title={`${email.attachments.length} attachment(s)`} style={{ marginRight: 4}}>📎</span>
                                )}
                                {email.subject}
                            </strong>
                            {!isRead && (
                                <span style={{
                                    width: 8,
                                    height: 8,
                                    borderRadius: "50%",
                                    background: "#007bff",
                                    display: "inline-block"
                                }}></span>
                            )}
                        </div>
                        <div style={{fontSize: 12, color: "#666"}}>
                            {previewBody.length > 100 ? `${previewBody.slice(0, 100)}...` : previewBody}
                        </div>
                    </div>
                );
            })}
        </div>
    );
}
