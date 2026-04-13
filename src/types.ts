
export interface User {
    email: string;
    teams: Team[]
}

export interface Team {
    email: string;
    alias?: string;  // Display name for the team (frontend-only)
    roles: Role[];
    tasks: Task[];
}

export interface Role {
    name: string;
    emails: Email[];
}

export interface Task {
    id: string;
    status: string;
    createdAt: Date;
    title: string;
    description?: string;
    due_time?: string | null;
}

export interface Attachment {
    filename: string;
    mimeType: string;
    size: number;
    attachmentId: string;
    gmailMessageId: string;
}

export interface Email {
    from: string;
    to: string;
    subject: string;
    body: string;
    receivedAt: string;
    role: string;
    read: boolean;
    unread?: boolean;  // Backend field — mapped to 'read' on fetch
    messageId: string;  // Required for threading replies
    gmailThreadId: string;  // Required for threading replies
    teamEmail?: string;  // Tracks which team this sent email belongs to
    attachments?: Attachment[];  // Email attachments (empty array if none)
}

export interface ScheduledEmail {
    id: string;
    fromEmail: string;
    recipients: string[];
    subject: string;
    body: string;
    sendAt: string;
    sent: boolean;
    cancelled: boolean;
    attachmentCount: number;
    attachments?: Attachment[];
}

