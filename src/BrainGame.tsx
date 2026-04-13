import {KebabMenu} from "./KebabMenu.tsx";
import {EmailList} from "./EmailList.tsx";
import {useCallback, useEffect, useRef, useState} from "react";
import JSZip from "jszip";
import {PptxVisualPreview} from "./PptxVisualPreview.tsx";
import type {EmailTemplate} from "./templates.ts";
import type {Attachment, Email, ScheduledEmail, Team, User} from "./types.ts";
import {bgiurl} from "./url.ts";

export function BrainGame() {
    const [teams, setTeams] = useState<Team[]>([]);
    const [users, setUsers] = useState<User[]>([]);
    const [selectedUser, setSelectedUser] = useState<string>('');
    const [selectedTeam, setSelectedTeam] = useState<string>('');
    const [selectedRole, setSelectedRole] = useState<string>('');
    const hasLoadedOnceRef = useRef(false);
    const teamsInFlightRef = useRef(false);
    const usersInFlightRef = useRef(false);
    const emailsInFlightRef = useRef(false);
    const addRoleInflightRef = useRef(false);
    const sentEmailsLoadedRef = useRef(false);
    const previousScheduledEmailsRef = useRef<ScheduledEmail[]>([]);
    const cancelledScheduledIdsRef = useRef<Set<string>>(new Set());
    const [showAddTeamModal, setShowAddTeamModal] = useState(false);
    const [showAddRoleModal, setShowAddRoleModal] = useState(false);
    const [showDeleteRoleModal, setShowDeleteRoleModal] = useState(false);
    const [activeManagementPanel, setActiveManagementPanel] = useState<'teams' | 'roles' | null>(null);
    const [newRoleName, setNewRoleName] = useState('');
    const [showRemoveTeamModal, setShowRemoveTeamModal] = useState(false);
    const [showCreateTeamModal, setShowCreateTeamModal] = useState(false);
    const [showCreateUserModal, setShowCreateUserModal] = useState(false);
    const [showCreateTaskModal, setShowCreateTaskModal] = useState(false);
    const [showEditTaskModal, setShowEditTaskModal] = useState(false);
    const [showDeleteTaskModal, setShowDeleteTaskModal] = useState(false);
    const [editingTaskId, setEditingTaskId] = useState<string | null>(null);
    const [deletingTaskId, setDeletingTaskId] = useState<string | null>(null);
    const [showEditAliasModal, setShowEditAliasModal] = useState(false);
    const [editingTeamEmail, setEditingTeamEmail] = useState<string>('');
    const [newAlias, setNewAlias] = useState('');
    const [newTeamEmail, setNewTeamEmail] = useState('');
    const [newUserEmail, setNewUserEmail] = useState('');
    const [newTaskTitle, setNewTaskTitle] = useState('');
    const [newTaskDescription, setNewTaskDescription] = useState('');
    const [newTaskDueTime, setNewTaskDueTime] = useState('');
    const [selectedEmail, setSelectedEmail] = useState<Email | null>(null);
    const [replySubject, setReplySubject] = useState('');
    const [replyBody, setReplyBody] = useState('');
    const [emailSent, setEmailSent] = useState(false);
    const [replySent, setReplySent] = useState(false);
    const [templateSearch, setTemplateSearch] = useState('');
    const [taskNotification, setTaskNotification] = useState<{ message: string; type: 'success' | 'error' } | null>(null);
    const [showNewEmailModal, setShowNewEmailModal] = useState(false);
    const [newEmailRecipients, setNewEmailRecipients] = useState<string[]>([]);
    const [showRecipientPicker, setShowRecipientPicker] = useState(false);
    const [newEmailSubject, setNewEmailSubject] = useState('');
    const [newEmailBody, setNewEmailBody] = useState('');
    const [sentEmails, setSentEmails] = useState<Email[]>([]);
    const [showMoveRoleDropdown, setShowMoveRoleDropdown] = useState(false);
    const [replyAttachments, setReplyAttachments] = useState<File[]>([]);
    const [newEmailAttachments, setNewEmailAttachments] = useState<File[]>([]);
    const [attachmentPreview, setAttachmentPreview] = useState<{
        name: string;
        mimeType: string;
        url: string;
        previewKind?: 'pptx-visual' | 'docx-text';
        arrayBuffer?: ArrayBuffer;
        parsedText?: string;
    } | null>(null);
    const [attachmentPreviewLoading, setAttachmentPreviewLoading] = useState(false);
    const [sessionNotification, setSessionNotification] = useState<{ message: string; type: 'success' | 'error' } | null>(null);
    const [scheduledEmails, setScheduledEmails] = useState<ScheduledEmail[]>([]);
    const [showScheduleModal, setShowScheduleModal] = useState(false);
    const [scheduleTime, setScheduleTime] = useState('');
    const [isSchedulingEmail, setIsSchedulingEmail] = useState(false);
    const scheduleInFlightRef = useRef(false);
    const attachmentBlobCacheRef = useRef<Map<string, Blob>>(new Map());
    const attachmentFetchInFlightRef = useRef<Map<string, Promise<Blob>>>(new Map());
    const localAttachmentBlobRef = useRef<Map<string, Blob>>(new Map());
    const [mailboxView, setMailboxView] = useState<'inbox' | 'outbox'>('inbox');
    const [emailTemplates, setEmailTemplates] = useState<EmailTemplate[]>([]);
    const [showCreateTemplateModal, setShowCreateTemplateModal] = useState(false);
    const [showEditTemplateModal, setShowEditTemplateModal] = useState(false);
    const [showDeleteTemplateModal, setShowDeleteTemplateModal] = useState(false);
    const [editingTemplate, setEditingTemplate] = useState<EmailTemplate | null>(null);
    const [deletingTemplate, setDeletingTemplate] = useState<EmailTemplate | null>(null);
    const [templateName, setTemplateName] = useState('');
    const [templateSubject, setTemplateSubject] = useState('');
    const [templateBody, setTemplateBody] = useState('');
    const [now, setNow] = useState(new Date());
    const [sessionStartTime, setSessionStartTime] = useState<string>(() => localStorage.getItem('sessionStartTime') || '');
    const [showSessionConfig, setShowSessionConfig] = useState(false);

    const showTaskNotification = (message: string, type: 'success' | 'error') => {
        setTaskNotification({ message, type });
        setTimeout(() => {
            setTaskNotification(null);
        }, 3000);
    };

    const showSessionNotification = (message: string, type: 'success' | 'error') => {
        setSessionNotification({ message, type });
        setTimeout(() => {
            setSessionNotification(null);
        }, 3000);
    };

    const formatHHmm = (date: Date): string => {
        const pad = (n: number) => String(n).padStart(2, '0');
        return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
    };

    const splitUrlAndTrailingPunctuation = (rawToken: string): { url: string; trailing: string } => {
        let url = rawToken;
        let trailing = '';
        while (/[.,!?;:)\]]$/.test(url)) {
            trailing = url.slice(-1) + trailing;
            url = url.slice(0, -1);
        }
        return { url, trailing };
    };

    const renderTextWithHyperlinks = (text: string) => {
        const urlRegex = /(https?:\/\/[^\s<>"']+)/gi;
        const parts = text.split(urlRegex);

        return parts.map((part, index) => {
            if (!/^https?:\/\//i.test(part)) {
                return <span key={`text-${index}`}>{part}</span>;
            }

            const { url, trailing } = splitUrlAndTrailingPunctuation(part);
            try {
                const parsed = new URL(url);
                if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
                    return <span key={`text-${index}`}>{part}</span>;
                }

                return (
                    <span key={`link-${index}`}>
                        <a
                            href={url}
                            target="_blank"
                            rel="noopener noreferrer"
                            style={{ color: '#1a73e8', textDecoration: 'underline' }}
                        >
                            {url}
                        </a>
                        {trailing}
                    </span>
                );
            } catch {
                return <span key={`text-${index}`}>{part}</span>;
            }
        });
    };

    const getActiveSessionStart = (date: Date, startTime: string): Date | null => {
        if (!startTime) return null;
        const [hoursRaw, minutesRaw] = startTime.split(':');
        const hours = Number(hoursRaw);
        const minutes = Number(minutesRaw);
        if (Number.isNaN(hours) || Number.isNaN(minutes)) return null;

        const durationMs = 90 * 60 * 1000;
        const todayStart = new Date(date);
        todayStart.setHours(hours, minutes, 0, 0);

        const yesterdayStart = new Date(todayStart);
        yesterdayStart.setDate(yesterdayStart.getDate() - 1);

        if (date.getTime() >= todayStart.getTime() && date.getTime() <= todayStart.getTime() + durationMs) {
            return todayStart;
        }

        if (date.getTime() >= yesterdayStart.getTime() && date.getTime() <= yesterdayStart.getTime() + durationMs) {
            return yesterdayStart;
        }

        return null;
    };

    const getSimulatedDateLabel = (date: Date, activeSessionStart: Date | null): string | null => {
        if (!activeSessionStart) return null;

        const elapsedMinutes = Math.floor((date.getTime() - activeSessionStart.getTime()) / (60 * 1000));
        const clampedDay = Math.max(0, Math.min(89, elapsedMinutes));

        const simulatedDate = new Date(2026, 0, 1);
        simulatedDate.setDate(simulatedDate.getDate() + clampedDay);
        const day = String(simulatedDate.getDate()).padStart(2, '0');
        const month = String(simulatedDate.getMonth() + 1).padStart(2, '0');
        return `${day}/${month}`;
    };

    const formatDayMonth = (date: Date): string => {
        const day = String(date.getDate()).padStart(2, '0');
        const month = String(date.getMonth() + 1).padStart(2, '0');
        return `${day}/${month}`;
    };

    const getTaskCountsForTeam = (team: Team, currentTime: Date) => {
        const tasks = team.tasks || [];
        const incompleteCount = tasks.filter(task => task.status !== 'COMPLETE').length;
        const overdueCount = tasks.filter(task => isTaskOverdue(task, currentTime)).length;
        return { incompleteCount, overdueCount };
    };

    const isTaskOverdue = (task: { status: string; due_time?: string | null }, currentTime: Date): boolean => {
        if (task.status === 'COMPLETE' || !task.due_time) return false;
        const due = new Date(task.due_time);
        if (Number.isNaN(due.getTime())) return false;
        return due.getTime() < currentTime.getTime();
    };

    const activeSessionStart = getActiveSessionStart(now, sessionStartTime);
    const simulatedDateLabel = getSimulatedDateLabel(now, activeSessionStart);
    const topDateLabel = simulatedDateLabel || formatDayMonth(now);
    const selectedEmailIsRead = selectedEmail ? (selectedEmail.read === true || selectedEmail.unread === false) : false;
    const canMarkSelectedEmailAsUnread = !!selectedEmail &&
        mailboxView === 'inbox' &&
        selectedRole !== 'Sent' &&
        selectedEmail.from === selectedTeam &&
        selectedEmailIsRead;
    const normalizedNewRoleName = newRoleName.trim().toLowerCase();
    const duplicateRole = teams
        .find(team => team.email === selectedTeam)
        ?.roles
        .find(role => role.name.trim().toLowerCase() === normalizedNewRoleName);
    const hasDuplicateRoleName = !!duplicateRole;
    const deletableRolesForSelectedTeam = (teams.find(team => team.email === selectedTeam)?.roles || [])
        .filter(role => role.name.trim().toLowerCase() !== 'unknown')
        .slice()
        .sort((a, b) => a.name.localeCompare(b.name));

    const toTimeLocal = (isoDate?: string | null): string => {
        if (!isoDate) return '';
        const date = new Date(isoDate);
        if (Number.isNaN(date.getTime())) return '';
        const pad = (n: number) => String(n).padStart(2, '0');
        return `${pad(date.getHours())}:${pad(date.getMinutes())}`;
    };

    const timeToTodayIso = (time: string): string | undefined => {
        if (!time) return undefined;
        const [hoursRaw, minutesRaw] = time.split(':');
        const hours = Number(hoursRaw);
        const minutes = Number(minutesRaw);

        if (Number.isNaN(hours) || Number.isNaN(minutes)) return undefined;

        const today = new Date();
        today.setHours(hours, minutes, 0, 0);
        return today.toISOString();
    };

    const canFetchAttachment = (att: Attachment): boolean => {
        const hasIds = !!att.attachmentId && !!att.gmailMessageId;
        if (!hasIds) return false;
        if (isLocalAttachment(att)) {
            return localAttachmentBlobRef.current.has(getAttachmentCacheKey(att));
        }
        return !!selectedUser;
    };

    const getAttachmentExtension = (filename?: string): string => {
        if (!filename) return '';
        const dotIndex = filename.lastIndexOf('.');
        if (dotIndex < 0) return '';
        return filename.slice(dotIndex + 1).toLowerCase();
    };

    const isAudioMime = (mimeType?: string, filename?: string): boolean => {
        const mime = (mimeType || '').toLowerCase();
        const name = (filename || '').toLowerCase();
        const ext = getAttachmentExtension(filename);
        return mime.startsWith('audio/') ||
            mime === 'audio/x-m4a' ||
            mime === 'audio/mp4' ||
            name.endsWith('.m4a') ||
            ['m4a', 'mp3', 'wav', 'ogg', 'aac'].includes(ext);
    };

    const isVideoMime = (mimeType?: string, filename?: string): boolean => {
        const mime = (mimeType || '').toLowerCase();
        const ext = getAttachmentExtension(filename);
        return mime.startsWith('video/') || ['mp4', 'mov', 'm4v', 'webm'].includes(ext);
    };

    const isOfficeAttachment = (mimeType?: string, filename?: string): boolean => {
        const mime = (mimeType || '').toLowerCase();
        const ext = getAttachmentExtension(filename);
        return ['doc', 'docx', 'ppt', 'pptx', 'pps', 'ppsx', 'xls', 'xlsx'].includes(ext) ||
            mime.includes('officedocument') ||
            mime.includes('msword') ||
            mime.includes('powerpoint') ||
            mime.includes('presentation') ||
            mime.includes('spreadsheet') ||
            mime.includes('excel');
    };

    const isPptxAttachment = (mimeType?: string, filename?: string): boolean => {
        const mime = (mimeType || '').toLowerCase();
        const ext = getAttachmentExtension(filename);
        return ext === 'pptx' || mime.includes('presentationml.presentation');
    };

    const isLocalAttachment = (att: Attachment): boolean => {
        return (att.attachmentId || '').startsWith('local-') || (att.gmailMessageId || '').startsWith('local-');
    };

    const isDocxAttachment = (mimeType?: string, filename?: string): boolean => {
        const mime = (mimeType || '').toLowerCase();
        const ext = getAttachmentExtension(filename);
        return ext === 'docx' || mime.includes('wordprocessingml.document');
    };

    const isPresentationAttachment = (att: Attachment): boolean => {
        const mime = (att.mimeType || '').toLowerCase();
        const name = (att.filename || '').toLowerCase();
        return mime.includes('presentation') || mime.includes('powerpoint') || name.endsWith('.ppt') || name.endsWith('.pptx');
    };

    const isPreviewableAttachment = (att: Attachment): boolean => {
        const mime = (att.mimeType || '').toLowerCase();
        const ext = getAttachmentExtension(att.filename);
        return mime.startsWith('image/') ||
            mime === 'application/pdf' ||
            mime.startsWith('text/') ||
            ['txt', 'md', 'csv', 'json', 'xml', 'log'].includes(ext) ||
            isAudioMime(att.mimeType, att.filename) ||
            isVideoMime(att.mimeType, att.filename) ||
            isOfficeAttachment(att.mimeType, att.filename);
    };

    const getAttachmentIcon = (att: Attachment): string => {
        const mime = (att.mimeType || '').toLowerCase();
        if (mime.startsWith('image/')) return '🖼️';
        if (mime === 'application/pdf') return '📄';
        if (isAudioMime(att.mimeType, att.filename)) return '🎵';
        if (isVideoMime(att.mimeType, att.filename)) return '🎬';
        if (isPresentationAttachment(att)) return '📽️';
        if (mime.includes('spreadsheet') || mime.includes('excel')) return '📊';
        if (mime.includes('word') || mime.includes('document')) return '📝';
        return '📁';
    };

    const getAttachmentDownloadUrl = (att: Attachment): string => {
        return `${bgiurl}/emails/${encodeURIComponent(selectedUser)}/attachments/${encodeURIComponent(att.gmailMessageId)}/${encodeURIComponent(att.attachmentId)}?filename=${encodeURIComponent(att.filename)}`;
    };

    const getAttachmentCacheKey = (att: Attachment): string => {
        if (isLocalAttachment(att)) {
            return `local::${att.gmailMessageId}::${att.attachmentId}`;
        }
        return `${selectedUser}::${att.gmailMessageId}::${att.attachmentId}`;
    };

    const createLocalOutgoingAttachments = (files: File[], localEmailId: string): Attachment[] => {
        return files.map((file, idx) => {
            const attachment: Attachment = {
                filename: file.name,
                mimeType: file.type || 'application/octet-stream',
                size: file.size,
                attachmentId: `local-${localEmailId}-${idx}`,
                gmailMessageId: `local-${localEmailId}`
            };
            localAttachmentBlobRef.current.set(getAttachmentCacheKey(attachment), file);
            return attachment;
        });
    };

    const fetchAttachmentBlob = async (att: Attachment): Promise<Blob> => {
        const cacheKey = getAttachmentCacheKey(att);
        const localBlob = localAttachmentBlobRef.current.get(cacheKey);
        if (localBlob) return localBlob;

        const cached = attachmentBlobCacheRef.current.get(cacheKey);
        if (cached) return cached;

        const inFlight = attachmentFetchInFlightRef.current.get(cacheKey);
        if (inFlight) return inFlight;

        const fetchPromise = (async () => {
        const res = await fetch(getAttachmentDownloadUrl(att));
        if (!res.ok) {
            const text = await res.text();
            throw new Error(text || `Failed to fetch attachment (${res.status})`);
        }

        const contentType = (res.headers.get('content-type') || '').toLowerCase();
        if (contentType.includes('application/json')) {
            const text = await res.text();
            throw new Error(text || 'Attachment endpoint returned JSON error instead of file bytes');
        }

        const originalBlob = await res.blob();
        if (originalBlob.size === 0) {
            throw new Error('Attachment file is empty');
        }

        // Backend often returns octet-stream; use attachment metadata for correct preview handling.
        const preferredType = att.mimeType || originalBlob.type;
        if (preferredType && originalBlob.type !== preferredType) {
            const normalizedBlob = new Blob([originalBlob], { type: preferredType });
            attachmentBlobCacheRef.current.set(cacheKey, normalizedBlob);
            return normalizedBlob;
        }

        attachmentBlobCacheRef.current.set(cacheKey, originalBlob);
        return originalBlob;
        })();

        attachmentFetchInFlightRef.current.set(cacheKey, fetchPromise);
        try {
            return await fetchPromise;
        } finally {
            attachmentFetchInFlightRef.current.delete(cacheKey);
        }
    };

    const prefetchEmailAttachments = (email: Email | null) => {
        if (!email?.attachments?.length) return;
        email.attachments.forEach(att => {
            if (!canFetchAttachment(att)) return;
            void fetchAttachmentBlob(att).catch(() => {
                // Ignore prefetch errors; on-demand preview/download will still show explicit errors.
            });
        });
    };

    const downloadAttachment = async (att: Attachment) => {
        if (!canFetchAttachment(att)) {
            showSessionNotification('Download not available for this attachment', 'error');
            return;
        }

        try {
            const blob = await fetchAttachmentBlob(att);
            const objectUrl = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = objectUrl;
            link.download = att.filename || 'attachment';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(objectUrl);
        } catch (err) {
            console.error('Failed to download attachment:', err);
            showSessionNotification(err instanceof Error ? err.message : 'Failed to download attachment', 'error');
        }
    };

    const clearAttachmentPreview = () => {
        setAttachmentPreview(prev => {
            if (prev?.url) URL.revokeObjectURL(prev.url);
            return null;
        });
    };

    const xmlToTextLines = (xml: string): string[] => {
        const doc = new DOMParser().parseFromString(xml, 'application/xml');
        const nodes = doc.querySelectorAll('a\\:t, w\\:t, t');
        const lines: string[] = [];
        nodes.forEach(node => {
            const text = (node.textContent || '').trim();
            if (text) lines.push(text);
        });
        return lines;
    };

    const extractDocxText = async (blob: Blob): Promise<string> => {
        const zip = await JSZip.loadAsync(await blob.arrayBuffer());
        const docXml = await zip.file('word/document.xml')?.async('string');
        if (!docXml) return '';
        return xmlToTextLines(docXml).join('\n');
    };

    const previewAttachment = async (att: Attachment) => {
        if (!canFetchAttachment(att)) {
            showSessionNotification('Preview not available for this attachment', 'error');
            return;
        }

        setAttachmentPreviewLoading(true);

        try {
            const blob = await fetchAttachmentBlob(att);
            const objectUrl = URL.createObjectURL(blob);
            const effectiveMime = blob.type || att.mimeType || '';

            if (isPptxAttachment(effectiveMime, att.filename)) {
                const arrayBuffer = await blob.arrayBuffer();
                setAttachmentPreview(prev => {
                    if (prev?.url) URL.revokeObjectURL(prev.url);
                    return {
                        name: att.filename,
                        mimeType: effectiveMime,
                        url: objectUrl,
                        previewKind: 'pptx-visual',
                        arrayBuffer
                    };
                });
                return;
            }

            if (isDocxAttachment(effectiveMime, att.filename)) {
                const text = await extractDocxText(blob);
                setAttachmentPreview(prev => {
                    if (prev?.url) URL.revokeObjectURL(prev.url);
                    return {
                        name: att.filename,
                        mimeType: effectiveMime,
                        url: objectUrl,
                        previewKind: 'docx-text',
                        parsedText: text
                    };
                });
                return;
            }

            setAttachmentPreview(prev => {
                if (prev?.url) URL.revokeObjectURL(prev.url);
                return { name: att.filename, mimeType: effectiveMime, url: objectUrl };
            });
        } catch (err) {
            console.error('Failed to preview attachment:', err);
            showSessionNotification('Failed to preview attachment', 'error');
        } finally {
            setAttachmentPreviewLoading(false);
        }
    };

    const fetchTemplates = async () => {
        try {
            const res = await fetch(`${bgiurl}/templates`);
            if (!res.ok) return;
            const data: EmailTemplate[] = await res.json();
            setEmailTemplates(data);
        } catch (err) {
            console.error('Failed to fetch templates:', err);
        }
    };

    const createTemplate = async () => {
        if (!templateName.trim() || !templateSubject.trim() || !templateBody.trim()) {
            alert('Please fill in all fields');
            return;
        }
        try {
            const res = await fetch(`${bgiurl}/templates`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name: templateName, subject: templateSubject, body: templateBody })
            });
            if (!res.ok) throw new Error(`Failed (${res.status})`);
            const uuid = await res.text();
            setEmailTemplates(prev => [...prev, { uuid, name: templateName, subject: templateSubject, body: templateBody }]);
            setShowCreateTemplateModal(false);
            setTemplateName('');
            setTemplateSubject('');
            setTemplateBody('');
            showSessionNotification('Template created', 'success');
        } catch (err) {
            console.error('Failed to create template:', err);
            showSessionNotification('Failed to create template', 'error');
        }
    };

    const updateTemplate = async () => {
        if (!editingTemplate || !templateName.trim() || !templateSubject.trim() || !templateBody.trim()) {
            alert('Please fill in all fields');
            return;
        }
        try {
            console.log(editingTemplate.uuid);
            const res = await fetch(`${bgiurl}/templates/${editingTemplate.uuid}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ name: templateName, subject: templateSubject, body: templateBody })
            });
            if (!res.ok) throw new Error(`Failed (${res.status})`);
            setEmailTemplates(prev => prev.map(t =>
                t.uuid === editingTemplate.uuid
                    ? { ...t, name: templateName, subject: templateSubject, body: templateBody }
                    : t
            ));
            setShowEditTemplateModal(false);
            setEditingTemplate(null);
            setTemplateName('');
            setTemplateSubject('');
            setTemplateBody('');
            showSessionNotification('Template updated', 'success');
        } catch (err) {
            console.error('Failed to update template:', err);
            showSessionNotification('Failed to update template', 'error');
        }
    };

    const deleteTemplate = async () => {
        if (!deletingTemplate) return;
        try {
            const res = await fetch(`${bgiurl}/templates/${deletingTemplate.uuid}`, {
                method: 'DELETE'
            });
            if (!res.ok) throw new Error(`Failed (${res.status})`);
            setEmailTemplates(prev => prev.filter(t => t.uuid !== deletingTemplate.uuid));
            setShowDeleteTemplateModal(false);
            setDeletingTemplate(null);
            showSessionNotification('Template deleted', 'success');
        } catch (err) {
            console.error('Failed to delete template:', err);
            showSessionNotification('Failed to delete template', 'error');
        }
    };


    const addTeamToUser = async (teamEmail: string) => {
        if (!selectedUser) {
            alert('Please select a user first');
            return;
        }

        try {
            const res = await fetch(`${bgiurl}/users/${selectedUser}/teams`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ teamEmail })
            });

            if (!res.ok) {
                throw new Error(`Failed to add team (${res.status})`);
            }

            // Refresh users to get updated team list
            await fetchUsers();
            setShowAddTeamModal(false);
        } catch (err) {
            console.error('Failed to add team to user:', err);
            alert('Failed to add team');
        }
    };

    const removeTeamFromUser = async (teamEmail: string) => {
        if (!selectedUser) {
            alert('Please select a user first');
            return;
        }

        try {
            const res = await fetch(`${bgiurl}/users/${selectedUser}/teams/${teamEmail}`, {
                method: 'DELETE'
            });

            if (!res.ok) {
                throw new Error(`Failed to remove team (${res.status})`);
            }

            await fetchUsers();
            setShowRemoveTeamModal(false);
        } catch (err) {
            console.error('Failed to remove team from user:', err);
            alert('Failed to remove team');
        }
    };

    const createTeam = async (teamEmail: string) => {
        if (!teamEmail.trim()) {
            showSessionNotification('Please enter a team email', 'error');
            return;
        }

        try {
            const res = await fetch(`${bgiurl}/teams`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: teamEmail })
            });

            if (!res.ok) {
                const errorMsg = await res.text();
                showSessionNotification(errorMsg || `Failed to create team (${res.status})`, 'error');
                return;
            }

            // Refresh teams to get the newly created team
            await fetchTeams();
            setNewTeamEmail('');
            setShowCreateTeamModal(false);
            showSessionNotification('Team created successfully!', 'success');
        } catch (err) {
            console.error('Failed to create team:', err);
            showSessionNotification('Failed to create team', 'error');
        }
    };

    const createUser = async (userEmail: string) => {
        if (!userEmail.trim()) {
            showSessionNotification('Please enter a user email', 'error');
            return;
        }

        try {
            const res = await fetch(`${bgiurl}/users`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: userEmail })
            });

            if (!res.ok) {
                const errorMsg = await res.text();
                showSessionNotification(errorMsg || `Failed to create user (${res.status})`, 'error');
                return;
            }

            // Refresh users to get the newly created user
            await fetchUsers();
            setNewUserEmail('');
            setShowCreateUserModal(false);
            showSessionNotification('User created successfully!', 'success');
        } catch (err) {
            console.error('Failed to create user:', err);
            showSessionNotification('Failed to create user', 'error');
        }
    };

    const createTask = async (teamEmail: string, title: string, description: string, dueTime: string) => {
        if (!title.trim()) {
            showTaskNotification('Please enter a task title', 'error');
            return;
        }

        const dueTimeIso = timeToTodayIso(dueTime);
        if (dueTime && !dueTimeIso) {
            showTaskNotification('Invalid due time', 'error');
            return;
        }

        try {
            console.log("Creating task with email " + teamEmail + " and title " + title + " and description " + description);
            const res = await fetch(`${bgiurl}/team/task`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    email: teamEmail,
                    title: title,
                    description: description.trim(),
                    ...(dueTimeIso ? { due_time: dueTimeIso } : {})
                })
            });

            if (!res.ok) {
                throw new Error(`Failed to create task (${res.status})`);
            }

            // Refresh teams to get the newly created task
            await fetchTeams();
            setNewTaskTitle('');
            setNewTaskDescription('');
            setNewTaskDueTime('');
            setShowCreateTaskModal(false);
            showTaskNotification('Task created successfully!', 'success');
        } catch (err) {
            console.error('Failed to create task:', err);
            showTaskNotification('Failed to create task', 'error');
        }
    };

    const deleteTask = async (teamEmail: string, taskId: string) => {
        try {
            const res = await fetch(`${bgiurl}/team/${teamEmail}/task/${taskId}`, {
                method: 'DELETE'
            });

            if (!res.ok) {
                throw new Error(`Failed to delete task (${res.status})`);
            }

            await fetchTeams();
            setShowDeleteTaskModal(false);
            setDeletingTaskId(null);
            showTaskNotification('Task deleted successfully!', 'success');
        } catch (err) {
            console.error('Failed to delete task:', err);
            showTaskNotification('Failed to delete task', 'error');
        }
    };

    const updateTask = async (teamEmail: string, taskId: string, title: string, description: string, dueTime: string) => {
        if (!title.trim()) {
            showTaskNotification('Please enter a task title', 'error');
            return;
        }

        const dueTimeIso = timeToTodayIso(dueTime);
        if (dueTime && !dueTimeIso) {
            showTaskNotification('Invalid due time', 'error');
            return;
        }

        try {
            const res = await fetch(`${bgiurl}/team/${teamEmail}/task/${taskId}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    title: title,
                    description: description.trim(),
                    ...(dueTimeIso ? { due_time: dueTimeIso } : { due_time: null })
                })
            });

            if (!res.ok) {
                throw new Error(`Failed to update task (${res.status})`);
            }

            await fetchTeams();
            setNewTaskTitle('');
            setNewTaskDescription('');
            setNewTaskDueTime('');
            setEditingTaskId(null);
            setShowEditTaskModal(false);
            showTaskNotification('Task updated successfully!', 'success');
        } catch (err) {
            console.error('Failed to update task:', err);
            showTaskNotification('Failed to update task', 'error');
        }
    };

    const markTaskComplete = async (teamEmail: string, taskId: string, currentStatus: string) => {
        const newStatus = currentStatus === 'COMPLETE' ? 'INCOMPLETE' : 'COMPLETE';

        try {
            const res = await fetch(`${bgiurl}/team/${teamEmail}/task/${taskId}/status`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ status: newStatus })
            });

            if (!res.ok) {
                throw new Error(`Failed to update task status (${res.status})`);
            }

            await fetchTeams();
            showTaskNotification(`Task marked as ${newStatus.toLowerCase()}`, 'success');
        } catch (err) {
            console.error('Failed to update task status:', err);
            showTaskNotification('Failed to update task status', 'error');
        }
    };

    const updateTeamAlias = (teamEmail: string, alias: string) => {
        // Update in local state
        setTeams(prev => prev.map(team =>
            team.email === teamEmail
                ? { ...team, alias: alias.trim() || undefined }
                : team
        ));

        // Persist to localStorage (frontend-only)
        const aliases = JSON.parse(localStorage.getItem('teamAliases') || '{}');
        if (alias.trim()) {
            aliases[teamEmail] = alias.trim();
        } else {
            delete aliases[teamEmail];
        }
        localStorage.setItem('teamAliases', JSON.stringify(aliases));

        setShowEditAliasModal(false);
        setNewAlias('');
        setEditingTeamEmail('');
    };

    const fetchUsers = useCallback(async () => {
        if (usersInFlightRef.current) return;
            usersInFlightRef.current = true;

        try {
            console.log("Sending request to users backend");
            const res = await fetch(`${bgiurl}/users`);
            if (!res.ok) {
                let msg = `Request failed (${res.status})`;
                try {
                    const errJson = await res.json();
                    msg = errJson?.message ?? msg;
                } catch {
                    throw new Error(msg);
                }
            }
            const data: User[] = await res.json();
            console.log("From Users backend: " + JSON.stringify(data));
            setUsers(data);

        } catch (err) {
            console.error('Failed to fetch users:', err);
        } finally {
            usersInFlightRef.current = false;
        }
    }, []);

    const fetchTeams = useCallback(async () => {
        if (teamsInFlightRef.current) return;
            teamsInFlightRef.current = true;

        try {
            const res = await fetch(`${bgiurl}/teams`);
            console.log("Sending request to teams backend");
            if (!res.ok) {
                let msg = `Request failed (${res.status})`;
                try {
                    const errJson = await res.json();
                    msg = errJson?.message ?? msg;
                } catch {
                    throw new Error(msg);
                }
            }
            const response = await res.json();
            console.log("From Teams backend: " + JSON.stringify(response));
            const data : Team[] = response['teams'];

            // Load aliases from localStorage
            const aliases = JSON.parse(localStorage.getItem('teamAliases') || '{}');
            const teamsWithAliases = data.map(team => ({
                ...team,
                alias: aliases[team.email] || team.alias
            }));

            // Merge with existing teams while preserving backend role-email changes.
            // Important: when a role is deleted backend moves its emails to Unknown;
            // we must not overwrite that with stale local role arrays.
            setTeams(prev => {
                if (prev.length === 0) return teamsWithAliases;

                const mergeEmails = (backendEmails: Email[] = [], localEmails: Email[] = []) => {
                    const combined = [...backendEmails, ...localEmails];
                    const deduped: Email[] = [];

                    for (const email of combined) {
                        const exists = deduped.some(existing =>
                            (existing.messageId && email.messageId && existing.messageId === email.messageId) ||
                            (
                                existing.from === email.from &&
                                existing.to === email.to &&
                                existing.subject === email.subject &&
                                existing.receivedAt === email.receivedAt
                            )
                        );
                        if (!exists) deduped.push(email);
                    }

                    return deduped;
                };

                return teamsWithAliases.map(newTeam => {
                    const existing = prev.find(t => t.email === newTeam.email);
                    if (!existing) return newTeam;
                    return {
                        ...newTeam,
                        roles: newTeam.roles.map(newRole => {
                            const existingRole = existing.roles.find(r => r.name === newRole.name);
                            return {
                                ...newRole,
                                emails: mergeEmails(newRole.emails || [], existingRole?.emails || [])
                            };
                        })
                    };
                });
            });

        } catch (err) {
            console.error('Failed to fetch teams:', err);
        } finally {
            teamsInFlightRef.current = false;
        }
    }, []);

    const fetchEmails = useCallback(async () => {
        if (!selectedUser || emailsInFlightRef.current) return;
            emailsInFlightRef.current = true;

        try {
            const params = new URLSearchParams({ userEmail: selectedUser });
            const res = await fetch(`${bgiurl}/users/emails?${params}`);
            console.log("Poking backend...")
            if (!res.ok) {
                let msg = `Request failed (${res.status})`;
                try {
                    const errJson = await res.json();
                    msg = errJson?.message ?? msg;
                } catch {
                    throw new Error(msg);
                }
            }

            const rawData: Email[] = await res.json();

            // Map backend's 'unread' field to frontend's 'read' field
            const data: Email[] = rawData.map(email => ({
                ...email,
                read: email.unread !== undefined ? !email.unread : (email.read ?? false),
            }));

            for (const email of data) {
                assignEmailToRole(email, email.from, email.role);
            }

            hasLoadedOnceRef.current = true;
        } catch (err) {
            console.error('Failed to fetch emails:', err);
        } finally {
            emailsInFlightRef.current = false;
        }
    }, [selectedUser]);

    const fetchSentEmails = useCallback(async () => {
        if (!selectedTeam) return;

        try {
            const res = await fetch(`${bgiurl}/emails/sent/${encodeURIComponent(selectedTeam)}`);
            if (!res.ok) {
                // If endpoint doesn't exist yet, just silently fail
                if (res.status === 404) {
                    console.log('Sent emails endpoint not yet implemented');
                    return;
                }
                throw new Error(`Failed to fetch sent emails (${res.status})`);
            }

            const data = await res.json();
            // Transform sent emails to have 'read' property
            const transformedData = data.map((email: Email) => ({
                ...email,
                read: true, // Sent emails are always "read"
                role: 'Sent', // Mark as Sent role
                teamEmail: email.teamEmail || selectedTeam // Track which team
            }));
            setSentEmails(prev => {
                const keyOf = (email: Email) => email.messageId || `${email.from}::${email.to}::${email.subject}::${email.receivedAt}`;
                const incomingKeys = new Set(transformedData.map(keyOf));
                const prevByKey = new Map(prev.map(email => [keyOf(email), email]));

                const mergedIncoming = transformedData.map((email: Email) => {
                    const key = keyOf(email);
                    const prevEmail = prevByKey.get(key);
                    if ((!email.attachments || email.attachments.length === 0) && prevEmail?.attachments?.length) {
                        return { ...email, attachments: prevEmail.attachments };
                    }
                    return email;
                });

                const localOnly = prev.filter(email => {
                    if (incomingKeys.has(keyOf(email))) return false;
                    return (email.attachments || []).some(att => isLocalAttachment(att));
                });

                return [...mergedIncoming, ...localOnly].sort((a, b) =>
                    new Date(b.receivedAt).getTime() - new Date(a.receivedAt).getTime()
                );
            });
        } catch (err) {
            console.error('Failed to fetch sent emails:', err);
        }
    }, [selectedTeam]);

    useEffect(() => {
        fetchUsers();
        fetchTeams();
        fetchTemplates();

        // Load sent emails from localStorage on startup
        try {
            const savedSentEmails = localStorage.getItem('sentEmails');
            if (savedSentEmails) {
                const parsedEmails = JSON.parse(savedSentEmails);
                setSentEmails(parsedEmails);
                console.log('✅ Loaded', parsedEmails.length, 'sent emails from localStorage');
            } else {
                console.log('ℹ️ No sent emails found in localStorage (first time)');
            }
        } catch (err) {
            console.error('❌ Failed to load sent emails from localStorage:', err);
        }
    }, []);

    useEffect(() => {
        console.log("Users state updated:", users);
    }, [users]);

    useEffect(() => {
        console.log("Teams state updated:", teams);
    }, [teams]);

    useEffect(() => {
        // Skip the very first render (when sentEmails is initialized as empty array)
        if (!sentEmailsLoadedRef.current) {
            sentEmailsLoadedRef.current = true;
            return;
        }

        // Save sent emails to localStorage whenever they change (after first render)
        try {
            localStorage.setItem('sentEmails', JSON.stringify(sentEmails));
            console.log('Saved', sentEmails.length, 'sent emails to localStorage');
        } catch (err) {
            console.error('Failed to save sent emails to localStorage:', err);
        }
    }, [sentEmails]);

    useEffect(() => {
        const id = window.setInterval(() => {
            setNow(new Date());
        }, 1000);

        return () => window.clearInterval(id);
    }, []);

    useEffect(() => {
        localStorage.setItem('sessionStartTime', sessionStartTime);
    }, [sessionStartTime]);

    useEffect(() => {
        // Keep cache scoped to the currently selected brain to avoid stale growth.
        attachmentBlobCacheRef.current.clear();
        attachmentFetchInFlightRef.current.clear();
        localAttachmentBlobRef.current.clear();
    }, [selectedUser]);

    useEffect(() => {
        prefetchEmailAttachments(selectedEmail);
    }, [selectedEmail, selectedUser]);

    useEffect(() => {
        // Clear selected email when switching roles or to Sent folder
        setSelectedEmail(null);
        setReplySubject('');
        setReplyBody('');
        setShowMoveRoleDropdown(false);
        clearAttachmentPreview();
    }, [selectedRole]);

    useEffect(() => {
        // Clear team/role/email selection when switching brains
        setSelectedTeam('');
        setSelectedRole('');
        setSelectedEmail(null);
        setMailboxView('inbox');
        setActiveManagementPanel(null);
        previousScheduledEmailsRef.current = [];
        cancelledScheduledIdsRef.current.clear();
    }, [selectedUser]);

    useEffect(() => {
        // Clear email view when switching teams
        setSelectedEmail(null);
        setReplySubject('');
        setReplyBody('');
        setShowMoveRoleDropdown(false);
        setActiveManagementPanel(null);
        clearAttachmentPreview();
    }, [selectedTeam]);

    useEffect(() => {
        clearAttachmentPreview();
        return () => {
            setAttachmentPreview(prev => {
                if (prev?.url) URL.revokeObjectURL(prev.url);
                return null;
            });
        };
    }, [selectedEmail]);

    useEffect(() => {
        fetchEmails();

        const id = window.setInterval(() => {
            fetchEmails();
        }, 5000);

        return () => {
            window.clearInterval(id);
        };
    }, [fetchEmails]);

    const addRoleToTeam = async (teamEmail: string, newRole: string) => {
        const trimmedRole = newRole.trim();
        if (!teamEmail || !trimmedRole) return;

        const existingTeam = teams.find(team => team.email === teamEmail);
        const duplicateRole = existingTeam?.roles.find(
            role => role.name.trim().toLowerCase() === trimmedRole.toLowerCase()
        );

        if (duplicateRole) {
            showSessionNotification(`Role "${duplicateRole.name}" already exists`, 'error');
            return;
        }

        setTeams(prev => prev.map(team =>
            team.email === teamEmail
                ? {
                    ...team,
                    roles: (team.roles.some(r => r.name.trim().toLowerCase() === trimmedRole.toLowerCase())
                        ? team.roles
                        : [...team.roles, {name: trimmedRole, emails: []}]
                    ).slice().sort((a, b) => a.name.localeCompare(b.name))
                }
                : team
        ));
        setNewRoleName('');
        setShowAddRoleModal(false);

        if (addRoleInflightRef.current) return;
            addRoleInflightRef.current = true;

        try {
            const res = await fetch(`${bgiurl}/teams/${teamEmail}/role?newRole=${encodeURIComponent(trimmedRole)}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' }
            });

            if (!res.ok) {
                let msg = `Request failed (${res.status})`;
                try {
                    const errJson = await res.json();
                    msg = errJson?.message ?? msg;
                } catch {
                    throw new Error(msg);
                }
            }

            hasLoadedOnceRef.current = true;
        } catch (err) {
            console.error('Failed to fetch emails:', err);
        } finally {
            addRoleInflightRef.current = false;
        }
    };

    const deleteRoleFromTeam = async (teamEmail: string, roleName: string) => {
        const trimmedRoleName = roleName.trim();
        if (!teamEmail || !trimmedRoleName) return;

        if (trimmedRoleName.toLowerCase() === 'unknown') {
            showSessionNotification('Unknown role cannot be deleted', 'error');
            return;
        }

        try {
            const res = await fetch(`${bgiurl}/teams/${encodeURIComponent(teamEmail)}/role?role=${encodeURIComponent(trimmedRoleName)}`, {
                method: 'DELETE'
            });

            const message = await res.text();
            if (!res.ok) {
                throw new Error(message || `Failed to delete role (${res.status})`);
            }

            if (selectedRole.toLowerCase() === trimmedRoleName.toLowerCase()) {
                setSelectedRole('');
                setSelectedEmail(null);
                setReplySubject('');
                setReplyBody('');
                setShowMoveRoleDropdown(false);
            }

            setShowDeleteRoleModal(false);
            showSessionNotification(message || `Deleted role ${trimmedRoleName}`, 'success');

            await fetchTeams();
            await fetchEmails();
        } catch (err) {
            console.error('Failed to delete role:', err);
            showSessionNotification(err instanceof Error ? err.message : 'Failed to delete role', 'error');
        }
    };

    //
    // const removeRoleFromTeam = (teamName: string, roleName: string) => {
    //     setTeams(prev => ({
    //         ...prev,
    //         [teamName]: prev[teamName].filter(role => role.name !== roleName)
    //     }));
    // };

    // When fetching emails, assign them to roles based on logic
    const assignEmailToRole = (email: Email, teamEmail: string, roleName: string) => {
        setTeams(prev => prev.map(team => {
            if (team.email !== teamEmail) return team;

            const existingRoles = team.roles || [];
            const hasRequestedRole = existingRoles.some(role => role.name === roleName);
            const unknownRole = existingRoles.find(role => role.name.toLowerCase() === 'unknown');
            const targetRoleName = hasRequestedRole ? roleName : (unknownRole?.name || 'Unknown');
            const roles = (hasRequestedRole || unknownRole)
                ? existingRoles
                : [...existingRoles, { name: 'Unknown', emails: [] }];

            return {
                ...team,
                roles: roles.map(role => {
                    if (role.name !== targetRoleName) return role;

                    const normalizedEmail: Email = {
                        ...email,
                        role: targetRoleName,
                        read: email.read ?? false,
                        unread: email.unread ?? true
                    };

                    const alreadyExists = (role.emails || []).some(e =>
                        e.messageId === email.messageId ||
                        (e.from === email.from &&
                            e.to === email.to &&
                            e.subject === email.subject &&
                            e.receivedAt === email.receivedAt)
                    );

                    return {
                        ...role,
                        emails: alreadyExists
                            ? (role.emails || []).map(e =>
                                (e.messageId === email.messageId ||
                                    (e.from === email.from &&
                                        e.to === email.to &&
                                        e.subject === email.subject &&
                                        e.receivedAt === email.receivedAt))
                                    ? {
                                        ...e,
                                        role: targetRoleName,
                                        read: e.read || normalizedEmail.read,
                                        unread: e.read ? false : (normalizedEmail.unread ?? e.unread)
                                    }
                                    : e
                            )
                            : [normalizedEmail, ...(role.emails || [])]
                    };
                })
            };
        }));
    };

    const sendReply = async () => {
        if (!selectedEmail) return;

        // Ensure "Re: " prefix is added if not already present
        const finalSubject = (replySubject || selectedEmail.subject).startsWith('Re: ')
            ? (replySubject || selectedEmail.subject)
            : `${replySubject || selectedEmail.subject}`;

        const sentEmailData = {
            from: selectedEmail.to,
            to: selectedEmail.from,
            subject: finalSubject,
            body: replyBody,
            sentAt: new Date().toISOString(),
            inReplyToMessageId: selectedEmail.messageId,
            gmailThreadId: selectedEmail.gmailThreadId
        };

        try {
            // Use the /reply endpoint for proper threading with FormData
            const formData = new FormData();
            formData.append('to', selectedEmail.from);
            formData.append('subject', finalSubject);
            formData.append('body', replyBody);
            formData.append('inReplyToMessageId', selectedEmail.messageId);
            formData.append('gmailThreadId', selectedEmail.gmailThreadId);
            for (const file of replyAttachments) {
                formData.append('attachments', file);
            }

            const res = await fetch(`${bgiurl}/emails/${selectedEmail.to}/reply`, {
                method: 'POST',
                body: formData
                // Do NOT set Content-Type header — browser sets it automatically with the boundary
            });

            if (!res.ok) {
                const errorText = await res.text();
                throw new Error(`Failed to send reply (${res.status}): ${errorText}`);
            }

            // Track sent email
            try {
                await fetch(`${bgiurl}/emails/sent`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(sentEmailData)
                });
            } catch (trackErr) {
                console.error('Failed to track sent email:', trackErr);
            }

            // Add to local sent emails immediately so it shows up in Sent folder
            // Build threaded body: reply at top, then quoted previous message
            const replyDate = new Date(selectedEmail.receivedAt);
            const dateStr = replyDate.toLocaleDateString('en-GB', {
                weekday: 'short',
                day: 'numeric',
                month: 'short',
                year: 'numeric'
            });
            const timeStr = replyDate.toLocaleTimeString('en-GB', {
                hour: '2-digit',
                minute: '2-digit'
            });
            const quotedPrevious = selectedEmail.body
                .split('\n')
                .map(line => `> ${line}`)
                .join('\n');
            const threadedBody = `${replyBody}\n\n> On ${dateStr} at ${timeStr}, ${selectedEmail.from} wrote:\n${quotedPrevious}`;

            const replyLocalId = `reply-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
            const sentAttachments = createLocalOutgoingAttachments(replyAttachments, replyLocalId);
            setSentEmails(prev => [{
                from: sentEmailData.from,
                to: sentEmailData.to,
                subject: sentEmailData.subject,
                body: threadedBody,
                receivedAt: sentEmailData.sentAt,
                role: 'Sent',
                read: true,
                messageId: `local-sent-reply-${replyLocalId}`,
                gmailThreadId: sentEmailData.gmailThreadId || '',
                teamEmail: selectedTeam,
                attachments: sentAttachments
            }, ...prev]);

            // Show success message only in the reply box (avoid duplicate global popup)
            setReplySent(true);
            setReplyBody('');
            setReplySubject('');
            setReplyAttachments([]);

            // Hide success message after 3 seconds
            setTimeout(() => {
                setReplySent(false);
            }, 3000);
        } catch (err) {
            console.error('Failed to send reply:', err);
            alert('Failed to send reply: ' + (err instanceof Error ? err.message : 'Unknown error'));
        }
    };

    const applyTemplate = (template: EmailTemplate) => {
        setReplySubject(template.subject);
        setReplyBody(template.body);
    };

    const applyTemplateToNewEmail = (template: EmailTemplate) => {
        setNewEmailSubject(template.subject);
        setNewEmailBody(template.body);
    };

    const sendNewEmail = async () => {
        console.log("Send new email called");
        if (!selectedUser || newEmailRecipients.length === 0 || !newEmailSubject.trim() || !newEmailBody.trim()) {
            alert('Please fill in all fields');
            return;
        }

        const sentAt = new Date().toISOString();

        try {
            // Send to each recipient individually
            for (const recipient of newEmailRecipients) {
                const formData = new FormData();
                formData.append('to', recipient);
                formData.append('subject', newEmailSubject);
                formData.append('body', newEmailBody);
                for (const file of newEmailAttachments) {
                    formData.append('attachments', file);
                }

                const res = await fetch(`${bgiurl}/emails/${selectedUser}/send`, {
                    method: 'POST',
                    body: formData
                });

                if (!res.ok) {
                    const errorText = await res.text();
                    throw new Error(`Failed to send email to ${recipient} (${res.status}): ${errorText}`);
                }

                // Track sent email
                try {
                    await fetch(`${bgiurl}/emails/sent`, {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            from: selectedUser,
                            to: recipient,
                            subject: newEmailSubject,
                            body: newEmailBody,
                            sentAt
                        })
                    });
                } catch (trackErr) {
                    console.error('Failed to track sent email:', trackErr);
                }

                // Add to local sent emails
                const sendLocalId = `send-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
                const sentAttachments = createLocalOutgoingAttachments(newEmailAttachments, sendLocalId);
                setSentEmails(prev => [{
                    from: selectedUser,
                    to: recipient,
                    subject: newEmailSubject,
                    body: newEmailBody,
                    receivedAt: sentAt,
                    role: 'Sent',
                    read: true,
                    messageId: `local-sent-${sendLocalId}`,
                    gmailThreadId: '',
                    teamEmail: selectedTeam,
                    attachments: sentAttachments
                }, ...prev]);
            }

            // Show success notification
            setEmailSent(true);

            // Clear form fields
            setNewEmailRecipients([]);
            setNewEmailSubject('');
            setNewEmailBody('');
            setNewEmailAttachments([]);

            // Close modal after showing success message
            setTimeout(() => {
                setShowNewEmailModal(false);
                setEmailSent(false);
            }, 800);
        } catch (err) {
            console.error('Failed to send email:', err);
            alert('Failed to send email: ' + (err instanceof Error ? err.message : 'Unknown error'));
        }
    };

    const scheduleEmail = async (sendAt: string) => {
        if (scheduleInFlightRef.current) return;

        if (!selectedUser || newEmailRecipients.length === 0 || !newEmailSubject.trim() || !newEmailBody.trim()) {
            alert('Please fill in all fields');
            return;
        }

        scheduleInFlightRef.current = true;
        setIsSchedulingEmail(true);

        try {
            const formData = new FormData();
            formData.append('recipients', newEmailRecipients.join(','));
            formData.append('subject', newEmailSubject);
            formData.append('body', newEmailBody);
            formData.append('sendAt', sendAt);
            for (const file of newEmailAttachments) {
                formData.append('attachments', file);
            }

            const res = await fetch(`${bgiurl}/emails/${selectedUser}/schedule`, {
                method: 'POST',
                body: formData
            });

            if (!res.ok) {
                const errorText = await res.text();
                throw new Error(`Failed to schedule email (${res.status}): ${errorText}`);
            }

            const data: ScheduledEmail = await res.json();
            const scheduledLocalId = `scheduled-${data.id}`;
            const scheduledAttachments = createLocalOutgoingAttachments(newEmailAttachments, scheduledLocalId);

            setScheduledEmails(prev => {
                const nextEmail = {
                    ...data,
                    attachmentCount: data.attachmentCount ?? scheduledAttachments.length,
                    attachments: scheduledAttachments
                };
                if (prev.some(item => item.id === nextEmail.id)) return prev;
                return [...prev, nextEmail];
            });

            setEmailSent(true);
            setNewEmailRecipients([]);
            setNewEmailSubject('');
            setNewEmailBody('');
            setNewEmailAttachments([]);
            setShowScheduleModal(false);
            setScheduleTime('');

            setTimeout(() => {
                setShowNewEmailModal(false);
                setEmailSent(false);
            }, 800);
        } catch (err) {
            console.error('Failed to schedule email:', err);
            alert('Failed to schedule email: ' + (err instanceof Error ? err.message : 'Unknown error'));
        } finally {
            scheduleInFlightRef.current = false;
            setIsSchedulingEmail(false);
        }
    };

    const cancelScheduledEmail = async (scheduledId: string) => {
        if (!selectedUser) return;

        try {
            const res = await fetch(`${bgiurl}/emails/${selectedUser}/schedule/${scheduledId}`, {
                method: 'DELETE'
            });

            if (!res.ok) {
                const errorText = await res.text();
                throw new Error(`Failed to cancel (${res.status}): ${errorText}`);
            }

            cancelledScheduledIdsRef.current.add(scheduledId);
            setScheduledEmails(prev => prev.filter(e => e.id !== scheduledId));
            showSessionNotification('Scheduled email cancelled', 'success');
        } catch (err) {
            console.error('Failed to cancel scheduled email:', err);
            showSessionNotification('Failed to cancel scheduled email', 'error');
        }
    };

    const fetchScheduledEmails = useCallback(async () => {
        if (!selectedUser) return;

        try {
            const res = await fetch(`${bgiurl}/emails/${selectedUser}/schedule`);
            if (!res.ok) return;
            const data: ScheduledEmail[] = await res.json();
            setScheduledEmails(prev => {
                const previousById = new Map(prev.map(item => [item.id, item]));
                return data.map(item => {
                    const previousItem = previousById.get(item.id);
                    return {
                        ...item,
                        attachmentCount: item.attachmentCount ?? previousItem?.attachmentCount ?? item.attachments?.length ?? 0,
                        attachments: item.attachments ?? previousItem?.attachments ?? []
                    };
                });
            });
        } catch (err) {
            console.error('Failed to fetch scheduled emails:', err);
        }
    }, [selectedUser]);

    useEffect(() => {
        if (!selectedUser) return;

        fetchScheduledEmails();
        const id = window.setInterval(() => {
            fetchScheduledEmails();
        }, 5000);

        return () => window.clearInterval(id);
    }, [fetchScheduledEmails, selectedUser]);

    useEffect(() => {
        const previousScheduled = previousScheduledEmailsRef.current;
        const currentIds = new Set(scheduledEmails.map(e => e.id));
        const removedScheduled = previousScheduled.filter(e => !currentIds.has(e.id));

        if (removedScheduled.length > 0 && selectedUser) {
            const sentScheduled = removedScheduled.filter(e => !cancelledScheduledIdsRef.current.has(e.id));

            // Clean up cancellation markers for any removed IDs.
            for (const removed of removedScheduled) {
                cancelledScheduledIdsRef.current.delete(removed.id);
            }

            if (sentScheduled.length > 0) {
                const promotedSentEmails: Email[] = sentScheduled.flatMap(scheduled =>
                    scheduled.recipients.map(recipient => ({
                        from: scheduled.fromEmail || selectedUser,
                        to: recipient,
                        subject: scheduled.subject,
                        body: scheduled.body,
                        receivedAt: scheduled.sendAt || new Date().toISOString(),
                        role: 'Sent',
                        read: true,
                        unread: false,
                        messageId: `scheduled-sent-${scheduled.id}-${recipient}`,
                        gmailThreadId: '',
                        teamEmail: recipient,
                        attachments: scheduled.attachments || []
                    }))
                );

                setSentEmails(prev => {
                    const existingIds = new Set(prev.map(email => email.messageId));
                    const uniquePromoted = promotedSentEmails.filter(email => !existingIds.has(email.messageId));
                    return uniquePromoted.length > 0 ? [...uniquePromoted, ...prev] : prev;
                });

                // Refresh backend-tracked sent emails when possible.
                fetchSentEmails();
            }
        }

        previousScheduledEmailsRef.current = scheduledEmails;
    }, [scheduledEmails, selectedUser, fetchSentEmails]);

    const markEmailAsUnread = async () => {
        if (!selectedEmail) return;

        // Update local state only (frontend-only feature)
        setTeams(prev => prev.map(team => ({
            ...team,
            roles: team.roles.map(role => ({
                ...role,
                emails: (role.emails || []).map(e =>
                    e.from === selectedEmail.from &&
                    e.to === selectedEmail.to &&
                    e.subject === selectedEmail.subject &&
                    e.receivedAt === selectedEmail.receivedAt
                        ? { ...e, read: false, unread: true }
                        : e
                )
            }))
        })));

        // Update selected email state
        setSelectedEmail({ ...selectedEmail, read: false, unread: true });

        // Persist unread state to backend
        if (selectedTeam && selectedEmail.messageId) {
            try {
                const res = await fetch(
                    `${bgiurl}/teams/${encodeURIComponent(selectedTeam)}/emails/unread?messageId=${encodeURIComponent(selectedEmail.messageId)}`,
                    { method: 'PUT' }
                );

                if (!res.ok) {
                    const errorText = await res.text();
                    throw new Error(errorText || `Failed to mark email unread (${res.status})`);
                }
            } catch (err) {
                console.error('Failed to mark email as unread:', err);
                showSessionNotification('Failed to mark email as unread', 'error');
            }
        }
    };

    const moveEmailToRole = async (newRole: string) => {
        if (!selectedEmail || !selectedTeam) return;

        try {
            const res = await fetch(`${bgiurl}/teams/email/role`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    messageId: selectedEmail.messageId,
                    teamEmail: selectedTeam,
                    oldRole: selectedEmail.role,
                    newRole: newRole
                })
            });

            if (!res.ok) {
                throw new Error(`Failed to move email (${res.status})`);
            }

            // Move email locally: remove from current role, add to new role
            const movedEmail = { ...selectedEmail, role: newRole };

            setTeams(prev => prev.map(team =>
                team.email === selectedTeam
                    ? {
                        ...team,
                        roles: team.roles.map(role => {
                            if (role.name === selectedRole) {
                                // Remove from current role
                                return {
                                    ...role,
                                    emails: (role.emails || []).filter(e =>
                                        !(e.messageId === selectedEmail.messageId)
                                    )
                                };
                            }
                            if (role.name === newRole) {
                                // Add to new role, sorted by receivedAt (newest first)
                                const updated = [movedEmail, ...(role.emails || [])];
                                updated.sort((a, b) => new Date(b.receivedAt).getTime() - new Date(a.receivedAt).getTime());
                                return {
                                    ...role,
                                    emails: updated
                                };
                            }
                            return role;
                        })
                    }
                    : team
            ));

            setSelectedEmail(null);
            setShowMoveRoleDropdown(false);
            showSessionNotification(`Email moved to ${newRole}`, 'success');
        } catch (err) {
            console.error('Failed to move email to role:', err);
            showSessionNotification('Failed to move email', 'error');
        }
    };


    // Get emails for selected role
    // const getSelectedRoleEmails = (): Email[] => {
    //     if (!selectedTeam || !selectedRole) return [];
    //     const team = teams[selectedTeam];
    //     const role = team?.find(r => r.name === selectedRole);
    //     return role?.emails || [];
    // };
        return (
            <div
                style={{
                    position: "fixed",
                    inset: 0,
                    padding: "3vh 4vw",
                    boxSizing: "border-box",
                    overflow: "hidden",
                    background: "#fff",

                    display: "flex",
                    flexDirection: "column",
                    gap: 12,
                }}
            >
                {/* Global Email Sent Notification */}
                {emailSent && (
                    <div style={{
                        position: 'fixed',
                        top: 92,
                        left: '50%',
                        transform: 'translateX(-50%)',
                        padding: '15px 30px',
                        background: '#28a745',
                        color: 'white',
                        borderRadius: 8,
                        boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
                        zIndex: 2000,
                        fontSize: 16,
                        fontWeight: 600,
                        animation: 'slideDown 0.3s ease-out'
                    }}>
                        ✓ Email Sent Successfully!
                    </div>
                )}

                {sessionNotification && (
                    <div style={{
                        position: 'fixed',
                        top: 92,
                        left: '50%',
                        transform: 'translateX(-50%)',
                        padding: '15px 30px',
                        background: sessionNotification.type === 'success' ? '#28a745' : '#dc3545',
                        color: 'white',
                        borderRadius: 8,
                        boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
                        zIndex: 2000,
                        fontSize: 16,
                        fontWeight: 600,
                    }}>
                        {sessionNotification.type === 'success' ? '✓' : '✗'} {sessionNotification.message}
                    </div>
                )}

                <div style={{
                    position: 'fixed',
                    top: 12,
                    left: '50%',
                    transform: 'translateX(-50%)',
                    background: '#ffffff',
                    border: '1px solid #d9dee8',
                    borderRadius: 10,
                    boxShadow: '0 3px 10px rgba(0,0,0,0.12)',
                    padding: '10px 14px',
                    zIndex: 1800,
                    display: 'flex',
                    alignItems: 'center',
                    gap: 10,
                    flexDirection: 'column'
                }}>
                    <div style={{display: 'flex', alignItems: 'center', gap: 8}}>
                        <div style={{fontSize: 14, fontWeight: 700}}>{formatHHmm(now)} | {topDateLabel}</div>
                        <button
                            onClick={() => setShowSessionConfig(prev => !prev)}
                            style={{
                                width: 24,
                                height: 24,
                                border: '1px solid #cbd5e1',
                                background: '#fff',
                                borderRadius: 4,
                                cursor: 'pointer',
                                lineHeight: '20px',
                                fontSize: 16,
                                color: '#475569',
                                padding: 0
                            }}
                            title="Session settings"
                        >
                            ...
                        </button>
                    </div>
                    {showSessionConfig && (
                        <div style={{display: 'flex', alignItems: 'center', gap: 8}}>
                            <input
                                type="time"
                                value={sessionStartTime}
                                onChange={(e) => setSessionStartTime(e.target.value)}
                                style={{padding: '4px 6px'}}
                            />
                            <button
                                onClick={() => {
                                    if (!sessionStartTime) {
                                        showSessionNotification('Please set a session start time', 'error');
                                        return;
                                    }
                                    showSessionNotification('Session time set', 'success');
                                    setShowSessionConfig(false);
                                }}
                                style={{
                                    padding: '4px 8px',
                                    border: '1px solid #1a73e8',
                                    background: '#1a73e8',
                                    color: 'white',
                                    borderRadius: 4,
                                    cursor: 'pointer',
                                    fontSize: 12
                                }}
                            >
                                Set Session Start Time
                            </button>
                        </div>
                    )}
                </div>

                <div style={{position: "absolute", top: 15, right: 15}}>
                    <KebabMenu
                        users={users}
                        selectedUser={selectedUser}
                        onSelectUser={setSelectedUser}
                        onCreateUser={() => setShowCreateUserModal(true)}
                        onCreateTeam={() => setShowCreateTeamModal(true)}
                        onStartSession={async () => {
                            if (!selectedUser) {
                                alert('Please select a user first');
                                return;
                            }
                            try {
                                const res = await fetch(`${bgiurl}/users/initialize`, {
                                    method: 'PUT',
                                    headers: {'Content-Type': 'application/json'},
                                    body: JSON.stringify({emailAddress: selectedUser})
                                });
                                if (!res.ok) throw new Error(`Failed (${res.status})`);
                                showSessionNotification('Session started for ' + selectedUser, 'success');
                            } catch (err) {
                                console.error('Failed to start session:', err);
                                showSessionNotification('Failed to start session', 'error');
                            }
                        }}
                        onEndSession={async () => {
                            if (!selectedUser) {
                                alert('Please select a user first');
                                return;
                            }
                            try {
                                const res = await fetch(`${bgiurl}/users/stop`, {
                                    method: 'PUT',
                                    headers: {'Content-Type': 'application/json'},
                                    body: JSON.stringify({emailAddress: selectedUser})
                                });
                                if (!res.ok) throw new Error(`Failed (${res.status})`);
                                localStorage.removeItem('sentEmails');
                                setSentEmails([]);
                                await fetchTeams();
                                await fetchUsers();
                                showSessionNotification('Session ended for ' + selectedUser, 'success');
                            } catch (err) {
                                console.error('Failed to end session:', err);
                                showSessionNotification('Failed to end session', 'error');
                            }
                        }}
                    />
                </div>

                <div style={{display: "flex", alignItems: "center", gap: 10}}>
                    <div style={{display: "flex", borderRadius: 6, overflow: "hidden", border: "1px solid #ccc"}}>
                        <button
                            onClick={() => {
                                setMailboxView('inbox');
                                setSelectedRole('');
                                setSelectedEmail(null);
                            }}
                            style={{
                                padding: "6px 16px",
                                fontSize: 13,
                                fontWeight: 600,
                                border: "none",
                                cursor: "pointer",
                                background: mailboxView === 'inbox' ? "#007bff" : "#f8f9fa",
                                color: mailboxView === 'inbox' ? "white" : "#333",
                            }}
                        >
                            📥 Inbox
                        </button>
                        <button
                            onClick={() => {
                                setMailboxView('outbox');
                                setSelectedRole('');
                                setSelectedEmail(null);
                            }}
                            style={{
                                padding: "6px 16px",
                                fontSize: 13,
                                fontWeight: 600,
                                border: "none",
                                borderLeft: "1px solid #ccc",
                                cursor: "pointer",
                                background: mailboxView === 'outbox' ? "#007bff" : "#f8f9fa",
                                color: mailboxView === 'outbox' ? "white" : "#333",
                            }}
                        >
                            📤 Outbox
                        </button>
                    </div>
                </div>
                <span>Current brain: {selectedUser === '' ? "none" : selectedUser}</span>
                {showCreateUserModal && (
                    <div style={{
                        position: 'fixed',
                        inset: 0,
                        background: 'rgba(0,0,0,0.5)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        zIndex: 1000
                    }}>
                        <div style={{
                            background: '#fff',
                            padding: 20,
                            borderRadius: 8,
                            maxWidth: 400,
                            zIndex: 1001
                        }}>
                            <h3>Create New Brain</h3>
                            <input
                                type="email"
                                value={newUserEmail}
                                onChange={(e) => setNewUserEmail(e.target.value)}
                                placeholder="user@example.com"
                                style={{padding: 8, width: '100%', marginBottom: 10}}
                            />
                            <button
                                onClick={() => createUser(newUserEmail)}
                                disabled={!newUserEmail.trim()}
                            >
                                Create
                            </button>
                            <button onClick={() => {
                                setShowCreateUserModal(false);
                                setNewUserEmail('');
                            }}>
                                Cancel
                            </button>
                        </div>
                    </div>
                )}

                {showCreateTeamModal && (
                    <div style={{
                        position: 'fixed',
                        inset: 0,
                        background: 'rgba(0,0,0,0.5)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        zIndex: 1000
                    }}>
                        <div style={{
                            background: '#fff',
                            padding: 20,
                            borderRadius: 8,
                            maxWidth: 400,
                            zIndex: 1001
                        }}>
                            <h3>Create New Team</h3>
                            <input
                                type="email"
                                value={newTeamEmail}
                                onChange={(e) => setNewTeamEmail(e.target.value)}
                                placeholder="team@example.com"
                                style={{padding: 8, width: '100%', marginBottom: 10}}
                            />
                            <button
                                onClick={() => createTeam(newTeamEmail)}
                                disabled={!newTeamEmail.trim()}
                            >
                                Create
                            </button>
                            <button onClick={() => {
                                setShowCreateTeamModal(false);
                                setNewTeamEmail('');
                            }}>
                                Cancel
                            </button>
                        </div>
                    </div>
                )}

                {teams.length > 0 && selectedUser && (
                    <div style={{display: "flex", flexWrap: "wrap", alignItems: "center"}}>
                        Teams:
                        {teams
                            .filter(team => {
                                const currentUser = users.find(u => u.email === selectedUser);
                                return currentUser?.teams?.find(userTeam => userTeam.email === team.email);
                            })
                            .sort((a, b) => a.email.localeCompare(b.email))
                            .map(team => {
                                const teamUnreadCount = (team.roles || []).reduce((total, role) => {
                                    const roleEmails = (role.emails || []).filter(email =>
                                        email.from === team.email && email.to === selectedUser
                                    );
                                    return total + roleEmails.filter(email => !(email.read || email.unread === false)).length;
                                }, 0);
                                const { incompleteCount, overdueCount } = getTaskCountsForTeam(team, now);

                                return (
                                <div key={team.email} style={{display: "flex", alignItems: "center", marginLeft: 8, marginRight: 12, position: "relative"}}>
                                    <button
                                        onClick={() => setSelectedTeam(team.email)}
                                        style={{
                                            padding: "8px 16px",
                                            background: selectedTeam === team.email ? "#007bff" : "#eee",
                                            color: selectedTeam === team.email ? "#fff" : "#000",
                                            border: "1px solid #ccc",
                                            borderRadius: "4px 0 0 4px",
                                            borderRight: "none",
                                            cursor: "pointer"
                                        }}
                                    >
                                        {team.alias || team.email}
                                    </button>
                                    {teamUnreadCount > 0 && (
                                        <span title="Unread emails" style={{
                                            position: "absolute",
                                            top: -6,
                                            left: -6,
                                            background: "#dc3545",
                                            color: "white",
                                            borderRadius: "50%",
                                            width: 18,
                                            height: 18,
                                            display: "flex",
                                            alignItems: "center",
                                            justifyContent: "center",
                                            fontSize: 10,
                                            fontWeight: 700,
                                            zIndex: 2,
                                            boxShadow: "0 1px 3px rgba(0,0,0,0.3)"
                                        }}>
                                            {teamUnreadCount > 99 ? '99+' : teamUnreadCount}
                                        </span>
                                    )}
                                    <button
                                        onClick={() => {
                                            setEditingTeamEmail(team.email);
                                            setNewAlias(team.alias || '');
                                            setShowEditAliasModal(true);
                                        }}
                                        style={{
                                            padding: "8px 8px",
                                            background: "#f8f9fa",
                                            border: "1px solid #ccc",
                                            borderRadius: "0 4px 4px 0",
                                            cursor: "pointer",
                                            fontSize: 12
                                        }}
                                        title="Edit alias"
                                    >
                                        ✏️
                                    </button>
                                    {(incompleteCount > 0 || overdueCount > 0) && (
                                        <span style={{
                                            position: "absolute",
                                            top: -6,
                                            left: teamUnreadCount > 0 ? 16 : -6,
                                            display: "flex",
                                            gap: 4,
                                            zIndex: 2
                                        }}>
                                            {incompleteCount > 0 && (
                                                <span title="Incomplete tasks" style={{
                                                    background: "#1a73e8",
                                                    color: "white",
                                                    borderRadius: "50%",
                                                    width: 18,
                                                    height: 18,
                                                    display: "flex",
                                                    alignItems: "center",
                                                    justifyContent: "center",
                                                    fontSize: 10,
                                                    fontWeight: 700,
                                                    boxShadow: "0 1px 3px rgba(0,0,0,0.3)"
                                                }}>
                                                    {incompleteCount > 99 ? '99+' : incompleteCount}
                                                </span>
                                            )}
                                            {overdueCount > 0 && (
                                                <span title="Overdue tasks" style={{
                                                    background: "#fbbc04",
                                                    color: "#000",
                                                    borderRadius: "50%",
                                                    width: 18,
                                                    height: 18,
                                                    display: "flex",
                                                    alignItems: "center",
                                                    justifyContent: "center",
                                                    fontSize: 10,
                                                    fontWeight: 700,
                                                    boxShadow: "0 1px 3px rgba(0,0,0,0.3)"
                                                }}>
                                                    {overdueCount > 99 ? '99+' : overdueCount}
                                                </span>
                                            )}
                                        </span>
                                    )}
                                </div>
                                );
                            })}
                        <div style={{display: 'flex', alignItems: 'center', marginLeft: 12}}>
                            <div style={{display: 'flex', gap: 8}}>
                                <button
                                    onClick={() => setActiveManagementPanel(prev => prev === 'teams' ? null : 'teams')}
                                    style={{
                                        background: activeManagementPanel === 'teams' ? '#1a73e8' : '#f8f9fa',
                                        color: activeManagementPanel === 'teams' ? 'white' : '#000',
                                        border: '1px solid #ccc',
                                        borderRadius: 4,
                                        padding: '8px 12px',
                                        cursor: 'pointer'
                                    }}
                                >
                                    Manage teams
                                </button>
                                <button
                                    onClick={() => setActiveManagementPanel(prev => prev === 'roles' ? null : 'roles')}
                                    disabled={!selectedTeam}
                                    style={{
                                        background: activeManagementPanel === 'roles' ? '#1a73e8' : '#f8f9fa',
                                        color: activeManagementPanel === 'roles' ? 'white' : '#000',
                                        border: '1px solid #ccc',
                                        borderRadius: 4,
                                        padding: '8px 12px',
                                        cursor: selectedTeam ? 'pointer' : 'not-allowed',
                                        opacity: selectedTeam ? 1 : 0.6
                                    }}
                                >
                                    Manage roles
                                </button>
                            </div>

                            {activeManagementPanel === 'teams' && (
                                <div style={{display: 'flex', gap: 8, marginLeft: 16}}>
                                    <button onClick={() => setShowAddTeamModal(true)}>Assign team</button>
                                    <button onClick={() => setShowRemoveTeamModal(true)}>Remove team</button>
                                </div>
                            )}

                            {activeManagementPanel === 'roles' && (
                                <div style={{display: 'flex', gap: 8, marginLeft: 16}}>
                                    <button onClick={() => setShowAddRoleModal(true)} disabled={!selectedTeam}>Add role</button>
                                    <button
                                        onClick={() => setShowDeleteRoleModal(true)}
                                        disabled={!selectedTeam || deletableRolesForSelectedTeam.length === 0}
                                    >
                                        Remove role
                                    </button>
                                </div>
                            )}
                        </div>
                            {showAddRoleModal && (
                                <div style={{
                                    position: 'fixed',
                                    inset: 0,
                                    background: 'rgba(0,0,0,0.5)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    zIndex: 1000
                                }}>
                                    <div style={{
                                        background: '#fff',
                                        padding: 20,
                                        borderRadius: 8,
                                        maxWidth: 400,
                                        zIndex: 1001
                                    }}>
                                        <h3>Add role to {selectedTeam}</h3>
                                        <input
                                            type="text"
                                            value={newRoleName}
                                            onChange={(e) => setNewRoleName(e.target.value)}
                                            placeholder="Role name"
                                            style={{padding: 8, width: '100%', marginBottom: 10}}
                                        />
                                        {normalizedNewRoleName && hasDuplicateRoleName && (
                                            <div style={{
                                                marginBottom: 10,
                                                color: '#dc3545',
                                                fontSize: 12,
                                                fontWeight: 600
                                            }}>
                                                Role "{duplicateRole?.name}" already exists (case-insensitive match).
                                            </div>
                                        )}
                                        <button
                                            onClick={() => addRoleToTeam(selectedTeam, newRoleName)}
                                            disabled={!newRoleName.trim() || hasDuplicateRoleName}
                                        >
                                            Add
                                        </button>
                                        <button onClick={() => {
                                            setShowAddRoleModal(false);
                                            setNewRoleName('');
                                        }}>
                                            Cancel
                                        </button>
                                    </div>
                                </div>
                            )}
                            {showDeleteRoleModal && (
                                <div style={{
                                    position: 'fixed',
                                    inset: 0,
                                    background: 'rgba(0,0,0,0.5)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    zIndex: 1000
                                }}>
                                    <div style={{
                                        background: '#fff',
                                        padding: 20,
                                        borderRadius: 8,
                                        maxWidth: 440,
                                        zIndex: 1001
                                    }}>
                                        <h3 style={{marginTop: 0}}>Remove role from {selectedTeam}</h3>
                                        <p style={{marginTop: 0, marginBottom: 10, color: '#555', fontSize: 13}}>
                                            Pick a role to remove. Emails in that role will be moved to <strong>Unknown</strong>.
                                        </p>

                                        {deletableRolesForSelectedTeam.length === 0 ? (
                                            <div style={{color: '#999', marginBottom: 12}}>No removable roles on this team.</div>
                                        ) : (
                                            <div style={{maxHeight: '45vh', overflowY: 'auto', marginBottom: 12}}>
                                                {deletableRolesForSelectedTeam.map(role => (
                                                    <button
                                                        key={role.name}
                                                        onClick={() => deleteRoleFromTeam(selectedTeam, role.name)}
                                                        style={{
                                                            display: 'block',
                                                            width: '100%',
                                                            padding: 10,
                                                            margin: '5px 0',
                                                            cursor: 'pointer',
                                                            textAlign: 'left'
                                                        }}
                                                    >
                                                        {role.name}
                                                    </button>
                                                ))}
                                            </div>
                                        )}

                                        <div style={{display: 'flex', gap: 8}}>
                                            <button
                                                onClick={() => setShowDeleteRoleModal(false)}
                                                style={{padding: '8px 12px'}}
                                            >
                                                Cancel
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            )}
                            {showAddTeamModal && (
                                <div style={{
                                    position: 'fixed',
                                    inset: 0,
                                    background: 'rgba(0,0,0,0.5)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    zIndex: 1000
                                }}>
                                    <div style={{
                                        background: '#fff',
                                        padding: 20,
                                        borderRadius: 8,
                                        maxWidth: 400,
                                        maxHeight: '50vh',
                                        display: 'flex',
                                        flexDirection: 'column',
                                        zIndex: 1001
                                    }}>
                                        <h3 style={{margin: '0 0 10px 0', flexShrink: 0}}>Add Team
                                            to {selectedUser}</h3>
                                        <div style={{overflowY: 'auto', flex: 1}}>
                                            {(() => {
                                                const assignedTeamEmails = new Set(
                                                    users.flatMap(u => (u.teams || []).map(t => t.email))
                                                );
                                                const unassignedTeams = teams.filter(t => !assignedTeamEmails.has(t.email))
                                                    .sort((a, b) => a.email.localeCompare(b.email));
                                                if (unassignedTeams.length === 0) {
                                                    return <p style={{color: '#999'}}>No unassigned teams available</p>;
                                                }
                                                return unassignedTeams.map(team => (
                                                    <button
                                                        key={team.email}
                                                        onClick={() => addTeamToUser(team.email)}
                                                        style={{
                                                            display: 'block',
                                                            width: '100%',
                                                            padding: 10,
                                                            margin: '5px 0',
                                                            cursor: 'pointer'
                                                        }}
                                                    >
                                                        {team.email}
                                                    </button>
                                                ));
                                            })()}
                                        </div>
                                        <button onClick={() => setShowAddTeamModal(false)}
                                                style={{marginTop: 10, flexShrink: 0}}>Cancel
                                        </button>
                                    </div>
                                </div>
                            )}

                            {showRemoveTeamModal && (
                                <div style={{
                                    position: 'fixed',
                                    inset: 0,
                                    background: 'rgba(0,0,0,0.5)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    zIndex: 1000
                                }}>
                                    <div style={{
                                        background: '#fff',
                                        padding: 20,
                                        borderRadius: 8,
                                        maxWidth: 400,
                                        zIndex: 1001
                                    }}>
                                        <h3>Remove Team from {selectedUser}</h3>
                                        {teams
                                            .filter(team => {
                                                const currentUser = users.find(u => u.email === selectedUser);
                                                return currentUser?.teams.find(userTeam => userTeam.email === team.email);
                                            })
                                            .sort((a, b) => a.email.localeCompare(b.email))
                                            .map(team => (
                                                <button
                                                    key={team.email}
                                                    onClick={() => removeTeamFromUser(team.email)}
                                                    style={{
                                                        display: 'block',
                                                        width: '100%',
                                                        padding: 10,
                                                        margin: '5px 0',
                                                        cursor: 'pointer'
                                                    }}
                                                >
                                                    {team.alias ? `${team.alias} (${team.email})` : team.email}
                                                </button>
                                            ))}
                                        <button onClick={() => setShowRemoveTeamModal(false)}>Cancel</button>
                                    </div>
                                </div>
                            )}

                            {/* Edit Alias Modal */}
                            {showEditAliasModal && (
                                <div style={{
                                    position: 'fixed',
                                    inset: 0,
                                    background: 'rgba(0,0,0,0.5)',
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center',
                                    zIndex: 1000
                                }}>
                                    <div style={{
                                        background: '#fff',
                                        padding: 20,
                                        borderRadius: 8,
                                        maxWidth: 400,
                                        width: '90%',
                                        zIndex: 1001
                                    }}>
                                        <h3>Edit Team Alias</h3>
                                        <p style={{fontSize: 14, color: "#666", marginBottom: 15}}>
                                            Team: {editingTeamEmail}
                                        </p>
                                        <input
                                            type="text"
                                            value={newAlias}
                                            onChange={(e) => setNewAlias(e.target.value)}
                                            placeholder="Enter team alias (e.g., The Best Team)"
                                            style={{
                                                padding: 8,
                                                width: '100%',
                                                marginBottom: 10,
                                                fontSize: 14,
                                                border: "1px solid #ccc",
                                                borderRadius: 4
                                            }}
                                        />
                                        <div style={{fontSize: 12, color: "#999", marginBottom: 15}}>
                                            Leave blank to use email address
                                        </div>
                                        <button
                                            onClick={() => updateTeamAlias(editingTeamEmail, newAlias)}
                                            style={{
                                                marginRight: 10,
                                                padding: "8px 16px",
                                                background: "#007bff",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: "pointer"
                                            }}
                                        >
                                            Save
                                        </button>
                                        <button
                                            onClick={() => {
                                                setShowEditAliasModal(false);
                                                setNewAlias('');
                                                setEditingTeamEmail('');
                                            }}
                                            style={{
                                                padding: "8px 16px",
                                                background: "#6c757d",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: "pointer"
                                            }}
                                        >
                                            Cancel
                                        </button>
                                    </div>
                                </div>
                            )}
                    </div>
                )}

                <div style={{position: "relative", height: "80vh", width: "100%"}}>
                    {/* Vertical role tabs on the left */}
                    {selectedUser && selectedTeam && (
                        <div style={{
                            position: "absolute",
                            left: "-50px",
                            top: 0,
                            bottom: 0,
                            width: "50px",
                            background: "#f8f9fa",
                            borderRadius: "8px 0 0 8px",
                            overflow: "auto",
                            boxShadow: "2px 0 8px rgba(0,0,0,0.1)",
                            zIndex: 1,
                            display: "flex",
                            flexDirection: "column"
                        }}>
                            {mailboxView === 'inbox' && (
                                <>
                                    {/* Regular roles */}
                                    {(teams.find(t => t.email === selectedTeam)?.roles || [])
                                        .slice()
                                        .sort((a, b) => a.name.localeCompare(b.name))
                                        .map((role) => {
                                        const filteredEmails = (role.emails || []).filter(email =>
                                            email.from === selectedTeam && email.to === selectedUser
                                        );
                                        const unreadCount = filteredEmails.filter(email => !(email.read || email.unread === false)).length;

                                        return (
                                            <div
                                                key={role.name}
                                                onClick={() => setSelectedRole(selectedRole === role.name ? '' : role.name)}
                                                style={{
                                                    cursor: "pointer",
                                                    padding: "20px 8px",
                                                    background: selectedRole === role.name ? "#007bff" : "white",
                                                    color: selectedRole === role.name ? "white" : "#000",
                                                    borderBottom: "1px solid #e0e0e0",
                                                    transition: "background 0.2s",
                                                    writingMode: "vertical-rl",
                                                    textOrientation: "mixed",
                                                    transform: "rotate(180deg)",
                                                    whiteSpace: "nowrap",
                                                    fontSize: 14,
                                                    fontWeight: 500,
                                                    textAlign: "center"
                                                }}
                                            >
                                                {role.name} ({unreadCount})
                                            </div>
                                        );
                                    })}
                                </>
                            )}

                            {mailboxView === 'outbox' && (
                                <>
                                    {/* Sent tab */}
                                    <div
                                        onClick={() => {
                                            setSelectedRole('Sent');
                                            fetchSentEmails();
                                        }}
                                        style={{
                                            cursor: "pointer",
                                            padding: "20px 8px",
                                            background: selectedRole === 'Sent' ? "#28a745" : "#f8f9fa",
                                            color: selectedRole === 'Sent' ? "white" : "#000",
                                            borderBottom: "1px solid #e0e0e0",
                                            transition: "background 0.2s",
                                            writingMode: "vertical-rl",
                                            textOrientation: "mixed",
                                            transform: "rotate(180deg)",
                                            whiteSpace: "nowrap",
                                            fontSize: 14,
                                            fontWeight: 500,
                                            textAlign: "center"
                                        }}
                                    >
                                        📤 Sent
                                        ({sentEmails.filter(e => e.teamEmail === selectedTeam && e.from === selectedUser).length})
                                    </div>

                                    {/* Scheduled tab */}
                                    <div
                                        onClick={() => {
                                            setSelectedRole('Scheduled');
                                            fetchScheduledEmails();
                                        }}
                                        style={{
                                            cursor: "pointer",
                                            padding: "20px 8px",
                                            background: selectedRole === 'Scheduled' ? "#f0ad4e" : "#f8f9fa",
                                            color: selectedRole === 'Scheduled' ? "white" : "#000",
                                            borderBottom: "1px solid #e0e0e0",
                                            transition: "background 0.2s",
                                            writingMode: "vertical-rl",
                                            textOrientation: "mixed",
                                            transform: "rotate(180deg)",
                                            whiteSpace: "nowrap",
                                            fontSize: 14,
                                            fontWeight: 500,
                                            textAlign: "center"
                                        }}
                                    >
                                        🕐 Scheduled ({scheduledEmails.length})
                                    </div>
                                </>
                            )}
                        </div>
                    )}

                    {/* Main grid with email list, email view, and templates */}
                    <div
                        style={{
                            display: "grid",
                            gridTemplateColumns: "30% 40% 30%",
                            minHeight: 0,
                            height: "100%",
                            width: "100%",
                            backgroundColor: "rgba(0, 0, 255, 0.1)",
                            borderRadius: 16
                        }}>
                        <div style={{background: "#eee", overflow: "auto"}}>
                            {selectedUser && selectedTeam && selectedRole && (() => {
                                // Check if "Sent" role is selected
                                if (selectedRole === 'Sent') {
                                    return (
                                        <EmailList
                                            emails={sentEmails.filter(e => e.teamEmail === selectedTeam && e.from === selectedUser)}
                                            onEmailClick={(email) => setSelectedEmail(email)}
                                            selectedEmail={selectedEmail}
                                        />
                                    );
                                }

                                // Check if "Scheduled" role is selected
                                if (selectedRole === 'Scheduled') {
                                    if (scheduledEmails.length === 0) {
                                        return (
                                            <div style={{
                                                display: "flex",
                                                alignItems: "center",
                                                justifyContent: "center",
                                                height: "100%",
                                                color: "#999"
                                            }}>
                                                No scheduled emails
                                            </div>
                                        );
                                    }
                                    return (
                                        <div style={{background: "#f9f9f9"}}>
                                            {scheduledEmails.map((se) => (
                                                <div
                                                    key={se.id}
                                                    onClick={() => setSelectedEmail({
                                                        from: se.fromEmail || selectedUser,
                                                        to: se.recipients.join(', '),
                                                        subject: se.subject,
                                                        body: se.body,
                                                        receivedAt: se.sendAt,
                                                        role: 'Scheduled',
                                                        read: true,
                                                        unread: false,
                                                        messageId: `scheduled-${se.id}`,
                                                        gmailThreadId: '',
                                                        teamEmail: selectedTeam,
                                                        attachments: se.attachments || []
                                                    })}
                                                    style={{
                                                        padding: 10,
                                                        borderBottom: "1px solid #ddd",
                                                        background: selectedEmail?.messageId === `scheduled-${se.id}` ? "#d4e9ff" : "white",
                                                        cursor: "pointer",
                                                        borderLeft: selectedEmail?.messageId === `scheduled-${se.id}` ? "3px solid #007bff" : "3px solid transparent"
                                                    }}
                                                >
                                                    <div style={{
                                                        display: "flex",
                                                        justifyContent: "space-between",
                                                        alignItems: "flex-start"
                                                    }}>
                                                        <div style={{flex: 1, minWidth: 0}}>
                                                            <strong style={{fontSize: 14}}>{se.subject}</strong>
                                                            <div style={{fontSize: 12, color: "#666", marginTop: 2}}>
                                                                To: {se.recipients.join(', ')}
                                                            </div>
                                                            <div style={{fontSize: 12, color: "#f0ad4e", marginTop: 2}}>
                                                                🕐 Scheduled for {new Date(se.sendAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                                                            </div>
                                                            {(se.attachmentCount || se.attachments?.length || 0) > 0 && (
                                                                <div
                                                                    style={{fontSize: 11, color: "#999", marginTop: 2}}>
                                                                    📎 {(se.attachmentCount || se.attachments?.length || 0)} attachment{(se.attachmentCount || se.attachments?.length || 0) > 1 ? 's' : ''}
                                                                </div>
                                                            )}
                                                        </div>
                                                        <button
                                                            onClick={(e) => {
                                                                e.stopPropagation();
                                                                cancelScheduledEmail(se.id);
                                                            }}
                                                            style={{
                                                                padding: "4px 10px",
                                                                fontSize: 12,
                                                                background: "#dc3545",
                                                                color: "white",
                                                                border: "none",
                                                                borderRadius: 4,
                                                                cursor: "pointer",
                                                                whiteSpace: "nowrap",
                                                                flexShrink: 0,
                                                                marginLeft: 8
                                                            }}
                                                        >
                                                            Cancel
                                                        </button>
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    );
                                }

                                // Otherwise show regular role emails
                                const role = teams.find(t => t.email === selectedTeam)?.roles.find(r => r.name === selectedRole);
                                const filteredEmails = (role?.emails || []).filter(email =>
                                    email.from === selectedTeam && email.to === selectedUser
                                );
                                return (
                                    <EmailList
                                        emails={filteredEmails}
                                        onEmailClick={(email) => {
                                            // Mark as read when clicked
                                            if (!(email.read || email.unread === false)) {
                                                const updatedEmail = {...email, read: true, unread: false};
                                                setSelectedEmail(updatedEmail);
                                                setTeams(prev => prev.map(team => ({
                                                    ...team,
                                                    roles: team.roles.map(role => ({
                                                        ...role,
                                                        emails: (role.emails || []).map(e =>
                                                            e.from === email.from &&
                                                            e.to === email.to &&
                                                            e.subject === email.subject &&
                                                            e.receivedAt === email.receivedAt
                                                                ? updatedEmail
                                                                : e
                                                        )
                                                    }))
                                                })));
                                                // Notify backend
                                                if (selectedTeam && email.messageId) {
                                                    fetch(
                                                        `${bgiurl}/teams/${encodeURIComponent(selectedTeam)}/emails/read?messageId=${encodeURIComponent(email.messageId)}`,
                                                        {method: 'PUT'}
                                                    ).catch(err => console.error('Failed to mark email as read:', err));
                                                }
                                            } else {
                                                setSelectedEmail(email);
                                            }
                                        }}
                                        selectedEmail={selectedEmail}
                                    />
                                );
                            })()}
                            {(!selectedUser || !selectedTeam || !selectedRole) && (
                                <div style={{
                                    display: "flex",
                                    alignItems: "center",
                                    justifyContent: "center",
                                    height: "100%",
                                    color: "#999"
                                }}>
                                    Select a role to view emails
                                </div>
                            )}
                        </div>
                        {selectedEmail ? (
                            <div style={{padding: 20, overflowY: "auto"}}>
                                <div style={{
                                    display: "flex",
                                    justifyContent: "space-between",
                                    alignItems: "flex-start",
                                    marginBottom: 10
                                }}>
                                    <h3 style={{margin: 0}}>{selectedEmail.subject}</h3>
                                    <div style={{display: "flex", gap: 8, alignItems: "center"}}>
                                        {selectedRole !== 'Sent' && selectedRole !== 'Scheduled' && (
                                            <div style={{position: "relative"}}>
                                                <button
                                                    onClick={() => setShowMoveRoleDropdown(!showMoveRoleDropdown)}
                                                    style={{
                                                        padding: "6px 12px",
                                                        fontSize: 12,
                                                        background: "#17a2b8",
                                                        color: "white",
                                                        border: "none",
                                                        borderRadius: 4,
                                                        cursor: "pointer",
                                                        whiteSpace: "nowrap"
                                                    }}
                                                >
                                                    Move to Role ▾
                                                </button>
                                                {showMoveRoleDropdown && (() => {
                                                    const currentTeam = teams.find(t => t.email === selectedTeam);
                                                    const availableRoles = (currentTeam?.roles || []).filter(r => r.name !== selectedRole);
                                                    return (
                                                        <div style={{
                                                            position: "absolute",
                                                            top: "100%",
                                                            right: 0,
                                                            marginTop: 4,
                                                            background: "#fff",
                                                            border: "1px solid #ccc",
                                                            borderRadius: 4,
                                                            boxShadow: "0 4px 12px rgba(0,0,0,0.15)",
                                                            zIndex: 100,
                                                            minWidth: 150,
                                                            overflow: "hidden"
                                                        }}>
                                                            {availableRoles.length === 0 ? (
                                                                <div style={{
                                                                    padding: "8px 12px",
                                                                    color: "#999",
                                                                    fontSize: 13
                                                                }}>
                                                                    No other roles
                                                                </div>
                                                            ) : availableRoles.map(role => (
                                                                <div
                                                                    key={role.name}
                                                                    onClick={() => moveEmailToRole(role.name)}
                                                                    style={{
                                                                        padding: "8px 12px",
                                                                        cursor: "pointer",
                                                                        fontSize: 13,
                                                                        borderBottom: "1px solid #eee"
                                                                    }}
                                                                    onMouseEnter={e => (e.currentTarget.style.background = "#f0f0f0")}
                                                                    onMouseLeave={e => (e.currentTarget.style.background = "#fff")}
                                                                >
                                                                    {role.name}
                                                                </div>
                                                            ))}
                                                        </div>
                                                    );
                                                })()}
                                            </div>
                                        )}
                                        {canMarkSelectedEmailAsUnread && (
                                            <button
                                                onClick={markEmailAsUnread}
                                                style={{
                                                    padding: "6px 12px",
                                                    fontSize: 12,
                                                    background: "#6c757d",
                                                    color: "white",
                                                    border: "none",
                                                    borderRadius: 4,
                                                    cursor: "pointer",
                                                    whiteSpace: "nowrap"
                                                }}
                                            >
                                                Mark as Unread
                                            </button>
                                        )}
                                    </div>
                                </div>
                                <p><strong>From:</strong> {(() => {
                                    const matchingTeam = teams.find(t => t.email === selectedEmail.from);
                                    return matchingTeam?.alias
                                        ? `${matchingTeam.alias} (${selectedEmail.from})`
                                        : selectedEmail.from;
                                })()}</p>
                                {(selectedRole === 'Sent' || selectedRole === 'Scheduled') && (
                                    <p><strong>To:</strong> {(() => {
                                        if (selectedRole === 'Scheduled') {
                                            return selectedEmail.to;
                                        }
                                        const matchingTeam = teams.find(t => t.email === selectedEmail.to);
                                        return matchingTeam?.alias
                                            ? `${matchingTeam.alias} (${selectedEmail.to})`
                                            : selectedEmail.to;
                                    })()}</p>
                                )}
                                <p><strong>At:</strong> {new Date(selectedEmail.receivedAt).toLocaleString()}</p>
                                <div style={{marginTop: 20, whiteSpace: "pre-wrap"}}>
                                    {renderTextWithHyperlinks(selectedEmail.body)}
                                </div>

                                {/* Attachments */}
                                {selectedEmail.attachments && selectedEmail.attachments.length > 0 && (
                                    <div style={{marginTop: 20, borderTop: "1px solid #ddd", paddingTop: 15}}>
                                        <h4 style={{margin: "0 0 10px 0", fontSize: 14, color: "#444"}}>
                                            📎 Attachments ({selectedEmail.attachments.length})
                                        </h4>
                                        {selectedEmail.attachments.map((att, idx) => (
                                            <div key={idx} style={{
                                                display: "flex",
                                                alignItems: "center",
                                                gap: 10,
                                                padding: "8px 12px",
                                                marginBottom: 6,
                                                background: "#f8f9fa",
                                                border: "1px solid #e0e0e0",
                                                borderRadius: 4
                                            }}>
                                            <span style={{fontSize: 18}}>{getAttachmentIcon(att)}</span>
                                                <div style={{flex: 1, minWidth: 0}}>
                                                    <div style={{
                                                        fontWeight: 500,
                                                        fontSize: 13,
                                                        overflow: "hidden",
                                                        textOverflow: "ellipsis",
                                                        whiteSpace: "nowrap"
                                                    }}>
                                                        {att.filename}
                                                    </div>
                                                    <div style={{fontSize: 11, color: "#999"}}>
                                                        {att.size < 1024
                                                            ? `${att.size} B`
                                                            : att.size < 1048576
                                                                ? `${(att.size / 1024).toFixed(1)} KB`
                                                                : `${(att.size / 1048576).toFixed(1)} MB`}
                                                    </div>
                                                </div>
                                                <div style={{ display: "flex", gap: 6 }}>
                                                    {isPreviewableAttachment(att) && canFetchAttachment(att) && (
                                                        <button
                                                            onClick={() => previewAttachment(att)}
                                                            style={{
                                                                padding: "4px 10px",
                                                                fontSize: 12,
                                                                background: "#17a2b8",
                                                                color: "white",
                                                                border: "none",
                                                                borderRadius: 4,
                                                                cursor: "pointer",
                                                                whiteSpace: "nowrap"
                                                            }}
                                                        >
                                                            Preview
                                                        </button>
                                                    )}
                                                    {canFetchAttachment(att) && (
                                                        <button
                                                            onClick={() => downloadAttachment(att)}
                                                            style={{
                                                                padding: "4px 12px",
                                                                fontSize: 12,
                                                                background: "#1a73e8",
                                                                color: "white",
                                                                border: "none",
                                                                borderRadius: 4,
                                                                cursor: "pointer",
                                                                whiteSpace: "nowrap"
                                                            }}
                                                        >
                                                            Download
                                                        </button>
                                                    )}
                                                </div>
                                            </div>
                                        ))}

                                    </div>
                                )}

                                {/* Only show reply section if not viewing a sent email */}
                                {selectedRole !== 'Sent' && selectedRole !== 'Scheduled' && (
                                    <div style={{
                                        marginTop: 40,
                                        borderTop: "1px solid #ddd",
                                        paddingTop: 20
                                    }}>
                                        <h4>Reply</h4>
                                        {replySent && (
                                            <div style={{
                                                padding: 10,
                                                marginBottom: 10,
                                                background: "#d4edda",
                                                color: "#155724",
                                                border: "1px solid #c3e6cb",
                                                borderRadius: 4
                                            }}>
                                                ✓ Reply sent successfully!
                                            </div>
                                        )}
                                        <div style={{marginBottom: 10}}>
                                            <input
                                                type="text"
                                                placeholder="Subject"
                                                value={replySubject || `${selectedEmail.subject}`}
                                                onChange={(e) => setReplySubject(e.target.value)}
                                                style={{
                                                    width: "100%",
                                                    padding: 8,
                                                    fontSize: 14,
                                                    border: "1px solid #ccc",
                                                    borderRadius: 4
                                                }}
                                            />
                                        </div>
                                        <div style={{marginBottom: 10}}>
                    <textarea
                        placeholder="Write your reply..."
                        rows={6}
                        value={replyBody}
                        onChange={(e) => setReplyBody(e.target.value)}
                        style={{
                            width: "100%",
                            padding: 8,
                            fontSize: 14,
                            border: "1px solid #ccc",
                            borderRadius: 4,
                            fontFamily: "inherit",
                            resize: "vertical"
                        }}
                    />
                                        </div>
                                        <div style={{marginBottom: 10}}>
                                            <label style={{
                                                display: "inline-flex",
                                                alignItems: "center",
                                                gap: 6,
                                                cursor: "pointer",
                                                fontSize: 13,
                                                color: "#1a73e8",
                                                padding: "6px 14px",
                                                border: "1px solid #1a73e8",
                                                borderRadius: 4,
                                                background: "white"
                                            }}>
                                                📎 Attach files
                                                <input
                                                    type="file"
                                                    multiple
                                                    onChange={(e) => {
                                                        if (e.target.files) {
                                                            setReplyAttachments(prev => [...prev, ...Array.from(e.target.files!)]);
                                                        }
                                                    }}
                                                    style={{display: "none"}}
                                                />
                                            </label>
                                            {replyAttachments.length > 0 && (
                                                <div style={{marginTop: 6}}>
                                                    {replyAttachments.map((file, idx) => (
                                                        <div key={idx} style={{
                                                            display: "flex",
                                                            alignItems: "center",
                                                            gap: 8,
                                                            padding: "4px 8px",
                                                            marginBottom: 4,
                                                            background: "#f0f0f0",
                                                            borderRadius: 4,
                                                            fontSize: 12
                                                        }}>
                                                        <span style={{
                                                            flex: 1,
                                                            overflow: "hidden",
                                                            textOverflow: "ellipsis",
                                                            whiteSpace: "nowrap"
                                                        }}>
                                                            {file.name} ({(file.size / 1024).toFixed(1)} KB)
                                                        </span>
                                                            <button
                                                                onClick={() => setReplyAttachments(prev => prev.filter((_, i) => i !== idx))}
                                                                style={{
                                                                    background: "none",
                                                                    border: "none",
                                                                    cursor: "pointer",
                                                                    color: "#dc3545",
                                                                    fontSize: 14,
                                                                    padding: 0
                                                                }}
                                                            >✕
                                                            </button>
                                                        </div>
                                                    ))}
                                                </div>
                                            )}
                                        </div>
                                        <button
                                            onClick={sendReply}
                                            disabled={!replyBody.trim()}
                                            style={{
                                                padding: "8px 16px",
                                                background: !replyBody.trim() ? "#ccc" : "#1a73e8",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: !replyBody.trim() ? "not-allowed" : "pointer"
                                            }}>
                                            Send Reply
                                        </button>
                                    </div>
                                )}
                            </div>
                        ) : (
                            <div style={{
                                background: "#f5f5f5",
                                display: "flex",
                                alignItems: "center",
                                justifyContent: "center",
                                color: "#999"
                            }}>
                                Select an email to view
                            </div>
                        )}
                        <div style={{
                            background: "#eee",
                            display: "flex",
                            flexDirection: "column",
                            overflow: "hidden"
                        }}>
                            {/* New Email Button */}
                            <div style={{padding: "10px 15px", borderBottom: "1px solid #ccc"}}>
                                <button
                                    onClick={() => {
                                        setNewEmailRecipients(selectedTeam ? [selectedTeam] : []);
                                        setShowNewEmailModal(true);
                                    }}
                                    disabled={!selectedUser || !selectedTeam}
                                    style={{
                                        width: "100%",
                                        padding: "10px",
                                        background: (!selectedUser || !selectedTeam) ? "#ccc" : "#28a745",
                                        color: "white",
                                        border: "none",
                                        borderRadius: 4,
                                        cursor: (!selectedUser || !selectedTeam) ? "not-allowed" : "pointer",
                                        fontSize: 14,
                                        fontWeight: 600
                                    }}
                                >
                                    ✉️ New Email
                                </button>
                            </div>

                            {/* Templates Section - Top Half */}
                            <div style={{
                                flex: "1",
                                padding: 15,
                                overflowY: "auto",
                                borderBottom: "2px solid #ccc"
                            }}>
                                <div style={{
                                    display: "flex",
                                    justifyContent: "space-between",
                                    alignItems: "center",
                                    marginBottom: 15
                                }}>
                                    <h4 style={{margin: 0}}>Templates</h4>
                                    <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                                        <input
                                            type="text"
                                            value={templateSearch}
                                            onChange={(e) => setTemplateSearch(e.target.value)}
                                            placeholder="Search..."
                                            style={{
                                                padding: "6px 10px",
                                                fontSize: 13,
                                                border: "1px solid #ccc",
                                                borderRadius: 4,
                                                width: "100px"
                                            }}
                                        />
                                        <button
                                            onClick={() => {
                                                setTemplateName('');
                                                setTemplateSubject('');
                                                setTemplateBody('');
                                                setShowCreateTemplateModal(true);
                                            }}
                                            style={{
                                                padding: "4px 10px",
                                                fontSize: 13,
                                                background: "#28a745",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: "pointer",
                                                whiteSpace: "nowrap"
                                            }}
                                        >
                                            + New
                                        </button>
                                    </div>
                                </div>
                                {(selectedEmail || showNewEmailModal) ? (
                                    <div>
                                        {emailTemplates
                                            .filter(template =>
                                                template.name.toLowerCase().includes(templateSearch.toLowerCase())
                                            )
                                            .map((template) => (
                                                <div
                                                    key={template.uuid}
                                                    style={{
                                                        display: "flex",
                                                        alignItems: "center",
                                                        marginBottom: 8,
                                                        background: "white",
                                                        border: "1px solid #ccc",
                                                        borderRadius: 4,
                                                        overflow: "hidden"
                                                    }}
                                                >
                                                    <button
                                                        onClick={() => showNewEmailModal ? applyTemplateToNewEmail(template) : applyTemplate(template)}
                                                        style={{
                                                            flex: 1,
                                                            padding: "10px 12px",
                                                            background: "transparent",
                                                            border: "none",
                                                            cursor: "pointer",
                                                            textAlign: "left",
                                                            transition: "background 0.2s",
                                                            minWidth: 0
                                                        }}
                                                        onMouseEnter={(e) => e.currentTarget.style.background = "#f0f0f0"}
                                                        onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}
                                                    >
                                                        <div style={{fontWeight: 600, marginBottom: 2, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"}}>
                                                            {template.name}
                                                        </div>
                                                        <div style={{fontSize: 12, color: "#666", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"}}>
                                                            {template.subject}
                                                        </div>
                                                    </button>
                                                    <div style={{
                                                        display: "flex",
                                                        flexDirection: "column",
                                                        gap: 4,
                                                        padding: "6px 8px",
                                                        flexShrink: 0
                                                    }}>
                                                        <button
                                                            onClick={() => {
                                                                setEditingTemplate(template);
                                                                setTemplateName(template.name);
                                                                setTemplateSubject(template.subject);
                                                                setTemplateBody(template.body);
                                                                setShowEditTemplateModal(true);
                                                            }}
                                                            style={{
                                                                padding: "2px 8px",
                                                                fontSize: 12,
                                                                background: "transparent",
                                                                border: "1px solid #007bff",
                                                                color: "#007bff",
                                                                borderRadius: 3,
                                                                cursor: "pointer"
                                                            }}
                                                        >
                                                            ✏️
                                                        </button>
                                                        <button
                                                            onClick={() => {
                                                                setDeletingTemplate(template);
                                                                setShowDeleteTemplateModal(true);
                                                            }}
                                                            style={{
                                                                padding: "2px 8px",
                                                                fontSize: 12,
                                                                background: "transparent",
                                                                border: "1px solid #dc3545",
                                                                color: "#dc3545",
                                                                borderRadius: 3,
                                                                cursor: "pointer"
                                                            }}
                                                        >
                                                            🗑
                                                        </button>
                                                    </div>
                                                </div>
                                            ))}
                                        {emailTemplates.filter(template =>
                                            template.name.toLowerCase().includes(templateSearch.toLowerCase())
                                        ).length === 0 && (
                                            <div style={{
                                                color: "#999",
                                                fontSize: 14,
                                                textAlign: "center",
                                                marginTop: 20
                                            }}>
                                                No templates found
                                            </div>
                                        )}
                                    </div>
                                ) : (
                                    <div>
                                        {emailTemplates
                                            .filter(template =>
                                                template.name.toLowerCase().includes(templateSearch.toLowerCase())
                                            )
                                            .map((template) => (
                                                <div
                                                    key={template.uuid}
                                                    style={{
                                                        display: "flex",
                                                        alignItems: "center",
                                                        marginBottom: 8,
                                                        background: "white",
                                                        border: "1px solid #ccc",
                                                        borderRadius: 4,
                                                        padding: "8px 12px"
                                                    }}
                                                >
                                                    <div style={{ flex: 1, minWidth: 0 }}>
                                                        <div style={{fontWeight: 600, marginBottom: 2, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"}}>
                                                            {template.name}
                                                        </div>
                                                        <div style={{fontSize: 12, color: "#666", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"}}>
                                                            {template.subject}
                                                        </div>
                                                    </div>
                                                    <div style={{
                                                        display: "flex",
                                                        flexDirection: "column",
                                                        gap: 4,
                                                        flexShrink: 0,
                                                        marginLeft: 8
                                                    }}>
                                                        <button
                                                            onClick={() => {
                                                                setEditingTemplate(template);
                                                                setTemplateName(template.name);
                                                                setTemplateSubject(template.subject);
                                                                setTemplateBody(template.body);
                                                                setShowEditTemplateModal(true);
                                                            }}
                                                            style={{
                                                                padding: "2px 8px",
                                                                fontSize: 12,
                                                                background: "transparent",
                                                                border: "1px solid #007bff",
                                                                color: "#007bff",
                                                                borderRadius: 3,
                                                                cursor: "pointer"
                                                            }}
                                                        >
                                                            ✏️
                                                        </button>
                                                        <button
                                                            onClick={() => {
                                                                setDeletingTemplate(template);
                                                                setShowDeleteTemplateModal(true);
                                                            }}
                                                            style={{
                                                                padding: "2px 8px",
                                                                fontSize: 12,
                                                                background: "transparent",
                                                                border: "1px solid #dc3545",
                                                                color: "#dc3545",
                                                                borderRadius: 3,
                                                                cursor: "pointer"
                                                            }}
                                                        >
                                                            🗑
                                                        </button>
                                                    </div>
                                                </div>
                                            ))}
                                        {emailTemplates.filter(template =>
                                            template.name.toLowerCase().includes(templateSearch.toLowerCase())
                                        ).length === 0 && (
                                            <div style={{ color: "#999", fontSize: 14, textAlign: "center", marginTop: 20 }}>
                                                No templates found
                                            </div>
                                        )}
                                    </div>
                                )}
                            </div>

                            {/* Tasks Section - Bottom Half */}
                            <div style={{
                                flex: "1",
                                padding: 15,
                                overflowY: "auto"
                            }}>
                                <div style={{
                                    display: "flex",
                                    justifyContent: "space-between",
                                    alignItems: "center",
                                    marginBottom: 15
                                }}>
                                    <h4 style={{margin: 0}}>Tasks</h4>
                                    {selectedTeam && (
                                        <button
                                            onClick={() => setShowCreateTaskModal(true)}
                                            style={{
                                                padding: "6px 12px",
                                                fontSize: 13,
                                                background: "#007bff",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: "pointer"
                                            }}
                                        >
                                            + New Task
                                        </button>
                                    )}
                                </div>

                                {taskNotification && (
                                    <div style={{
                                        padding: 10,
                                        marginBottom: 10,
                                        background: taskNotification.type === 'success' ? "#d4edda" : "#f8d7da",
                                        color: taskNotification.type === 'success' ? "#155724" : "#721c24",
                                        border: `1px solid ${taskNotification.type === 'success' ? "#c3e6cb" : "#f5c6cb"}`,
                                        borderRadius: 4,
                                        fontSize: 14
                                    }}>
                                        {taskNotification.type === 'success' ? '✓' : '✗'} {taskNotification.message}
                                    </div>
                                )}

                                {selectedTeam ? (
                                    <div>
                                        {(() => {
                                            const team = teams.find(t => t.email === selectedTeam);
                                            const tasks = [...(team?.tasks || [])].sort((a, b) =>
                                                new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime()
                                            );

                                            if (tasks.length === 0) {
                                                return (
                                                    <div style={{color: "#999", fontSize: 14}}>
                                                        No tasks for this team
                                                    </div>
                                                );
                                            }

                                            return tasks.map((task) => {
                                                const taskOverdue = isTaskOverdue(task, now);

                                                return (
                                                <div
                                                    key={task.id}
                                                    style={{
                                                        padding: 12,
                                                        marginBottom: 8,
                                                        background: task.status === 'COMPLETE' ? "#f0f0f0" : "white",
                                                        border: "1px solid #ccc",
                                                        borderRadius: 4,
                                                        opacity: task.status === 'COMPLETE' ? 0.6 : 1,
                                                        position: 'relative'
                                                    }}
                                                >
                                                    {taskOverdue && (
                                                        <span style={{
                                                            position: 'absolute',
                                                            top: 8,
                                                            right: 8,
                                                            color: '#dc3545',
                                                            fontSize: 12,
                                                            fontWeight: 700
                                                        }}>
                                                            OVERDUE
                                                        </span>
                                                    )}
                                                    <div style={{
                                                        fontSize: 15,
                                                        fontWeight: 600,
                                                        marginBottom: 6,
                                                        textDecoration: task.status === 'COMPLETE' ? 'line-through' : 'none'
                                                    }}>
                                                        {task.title}
                                                    </div>
                                                    <div style={{
                                                        fontSize: 13,
                                                        color: "#666",
                                                        lineHeight: 1.4,
                                                        textDecoration: task.status === 'COMPLETE' ? 'line-through' : 'none'
                                                    }}>
                                                        {task.description}
                                                    </div>
                                                    {task.due_time && (
                                                        <div style={{
                                                            fontSize: 12,
                                                            color: "#6c757d",
                                                            marginTop: 6,
                                                            textDecoration: task.status === 'COMPLETE' ? 'line-through' : 'none'
                                                        }}>
                                                            Due: {toTimeLocal(task.due_time)}
                                                        </div>
                                                    )}
                                                    <div style={{
                                                        fontSize: 11,
                                                        color: "#999",
                                                        marginTop: 6,
                                                        marginBottom: 8
                                                    }}>
                                                    </div>
                                                    <div style={{display: 'flex', gap: 6, flexWrap: 'wrap'}}>
                                                        <button
                                                            onClick={() => markTaskComplete(selectedTeam, task.id, task.status)}
                                                            style={{
                                                                padding: "4px 8px",
                                                                fontSize: 12,
                                                                background: task.status === 'COMPLETE' ? "#28a745" : "#6c757d",
                                                                color: "white",
                                                                border: "none",
                                                                borderRadius: 3,
                                                                cursor: "pointer"
                                                            }}
                                                        >
                                                            {task.status === 'COMPLETE' ? '✓ Complete' : 'Mark Complete'}
                                                        </button>
                                                        <button
                                                            onClick={() => {
                                                                setEditingTaskId(task.id);
                                                                setNewTaskTitle(task.title);
                                                                setNewTaskDescription(task.description || '');
                                                                setNewTaskDueTime(toTimeLocal(task.due_time));
                                                                setShowEditTaskModal(true);
                                                            }}
                                                            style={{
                                                                padding: "4px 8px",
                                                                fontSize: 12,
                                                                background: "#007bff",
                                                                color: "white",
                                                                border: "none",
                                                                borderRadius: 3,
                                                                cursor: "pointer"
                                                            }}
                                                        >
                                                            Edit
                                                        </button>
                                                        <button
                                                            onClick={() => {
                                                                setDeletingTaskId(task.id);
                                                                setShowDeleteTaskModal(true);
                                                            }}
                                                            style={{
                                                                padding: "4px 8px",
                                                                fontSize: 12,
                                                                background: "#dc3545",
                                                                color: "white",
                                                                border: "none",
                                                                borderRadius: 3,
                                                                cursor: "pointer"
                                                            }}
                                                        >
                                                            Delete
                                                        </button>
                                                    </div>
                                                </div>
                                            );
                                            });
                                        })()}
                                    </div>
                                ) : (
                                    <div style={{color: "#999", fontSize: 14}}>
                                        Select a team to view tasks
                                    </div>
                                )}
                            </div>
                        </div>

                        {/* Create Task Modal */}
                        {showCreateTaskModal && (
                            <div style={{
                                position: 'fixed',
                                inset: 0,
                                background: 'rgba(0,0,0,0.5)',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                zIndex: 1000
                            }}>
                                <div style={{
                                    background: '#fff',
                                    padding: 20,
                                    borderRadius: 8,
                                    maxWidth: 500,
                                    width: '90%',
                                    zIndex: 1001
                                }}>
                                    <h3>Create New Task for {selectedTeam}</h3>
                                    <input
                                        type="text"
                                        value={newTaskTitle}
                                        onChange={(e) => setNewTaskTitle(e.target.value)}
                                        placeholder="Task title"
                                        style={{padding: 8, width: '100%', marginBottom: 10}}
                                    />
                                    <textarea
                                        value={newTaskDescription}
                                        onChange={(e) => setNewTaskDescription(e.target.value)}
                                        placeholder="Task description (optional)"
                                        rows={4}
                                        style={{
                                            padding: 8,
                                            width: '100%',
                                            marginBottom: 10,
                                            fontFamily: 'inherit',
                                            resize: 'vertical'
                                        }}
                                    />
                                    <input
                                        type="time"
                                        value={newTaskDueTime}
                                        onChange={(e) => setNewTaskDueTime(e.target.value)}
                                        style={{padding: 8, width: '100%', marginBottom: 10}}
                                    />
                                    <button
                                        onClick={() => createTask(selectedTeam, newTaskTitle, newTaskDescription, newTaskDueTime)}
                                        disabled={!newTaskTitle.trim()}
                                        style={{
                                            marginRight: 10,
                                            padding: "8px 16px",
                                            background: (!newTaskTitle.trim()) ? "#ccc" : "#007bff",
                                            color: "white",
                                            border: "none",
                                            borderRadius: 4,
                                            cursor: (!newTaskTitle.trim()) ? "not-allowed" : "pointer"
                                        }}
                                    >
                                        Create
                                    </button>
                                    <button
                                        onClick={() => {
                                            setShowCreateTaskModal(false);
                                            setNewTaskTitle('');
                                            setNewTaskDescription('');
                                            setNewTaskDueTime('');
                                        }}
                                        style={{
                                            padding: "8px 16px",
                                            background: "#6c757d",
                                            color: "white",
                                            border: "none",
                                            borderRadius: 4,
                                            cursor: "pointer"
                                        }}
                                    >
                                        Cancel
                                    </button>
                                </div>
                            </div>
                        )}

                        {/* Edit Task Modal */}
                        {showEditTaskModal && editingTaskId && (
                            <div style={{
                                position: 'fixed',
                                inset: 0,
                                background: 'rgba(0,0,0,0.5)',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                zIndex: 1000
                            }}>
                                <div style={{
                                    background: '#fff',
                                    padding: 20,
                                    borderRadius: 8,
                                    maxWidth: 500,
                                    width: '90%',
                                    zIndex: 1001
                                }}>
                                    <h3>Edit Task</h3>
                                    <input
                                        type="text"
                                        value={newTaskTitle}
                                        onChange={(e) => setNewTaskTitle(e.target.value)}
                                        placeholder="Task title"
                                        style={{padding: 8, width: '100%', marginBottom: 10}}
                                    />
                                    <textarea
                                        value={newTaskDescription}
                                        onChange={(e) => setNewTaskDescription(e.target.value)}
                                        placeholder="Task description (optional)"
                                        rows={4}
                                        style={{
                                            padding: 8,
                                            width: '100%',
                                            marginBottom: 10,
                                            fontFamily: 'inherit',
                                            resize: 'vertical'
                                        }}
                                    />
                                    <input
                                        type="time"
                                        value={newTaskDueTime}
                                        onChange={(e) => setNewTaskDueTime(e.target.value)}
                                        style={{padding: 8, width: '100%', marginBottom: 10}}
                                    />
                                    <button
                                        onClick={() => updateTask(selectedTeam, editingTaskId, newTaskTitle, newTaskDescription, newTaskDueTime)}
                                        disabled={!newTaskTitle.trim()}
                                        style={{
                                            marginRight: 10,
                                            padding: "8px 16px",
                                            background: (!newTaskTitle.trim()) ? "#ccc" : "#007bff",
                                            color: "white",
                                            border: "none",
                                            borderRadius: 4,
                                            cursor: (!newTaskTitle.trim()) ? "not-allowed" : "pointer"
                                        }}
                                    >
                                        Update
                                    </button>
                                    <button
                                        onClick={() => {
                                            setShowEditTaskModal(false);
                                            setNewTaskTitle('');
                                            setNewTaskDescription('');
                                            setNewTaskDueTime('');
                                            setEditingTaskId(null);
                                        }}
                                        style={{
                                            padding: "8px 16px",
                                            background: "#6c757d",
                                            color: "white",
                                            border: "none",
                                            borderRadius: 4,
                                            cursor: "pointer"
                                        }}
                                    >
                                        Cancel
                                    </button>
                                </div>
                            </div>
                        )}

                        {/* Delete Task Confirmation Modal */}
                        {showDeleteTaskModal && deletingTaskId && (
                            <div style={{
                                position: 'fixed',
                                inset: 0,
                                background: 'rgba(0,0,0,0.5)',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                zIndex: 1000
                            }}>
                                <div style={{
                                    background: '#fff',
                                    padding: 20,
                                    borderRadius: 8,
                                    maxWidth: 400,
                                    width: '90%',
                                    zIndex: 1001
                                }}>
                                    <h3 style={{marginTop: 0}}>Delete Task</h3>
                                    <p>Are you sure you want to delete this task? This action cannot be undone.</p>
                                    <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 20 }}>
                                        <button
                                            onClick={() => {
                                                setShowDeleteTaskModal(false);
                                                setDeletingTaskId(null);
                                            }}
                                            style={{
                                                padding: "8px 16px",
                                                background: "#6c757d",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: "pointer"
                                            }}
                                        >
                                            Cancel
                                        </button>
                                        <button
                                            onClick={() => deleteTask(selectedTeam, deletingTaskId)}
                                            style={{
                                                padding: "8px 16px",
                                                background: "#dc3545",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: "pointer"
                                            }}
                                        >
                                            Delete
                                        </button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {/* New Email Modal */}
                        {showNewEmailModal && (
                            <div style={{
                                position: 'fixed',
                                top: 0,
                                left: 0,
                                right: '35%', // Leave space for templates on the right
                                bottom: 0,
                                background: 'rgba(0,0,0,0.3)',
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'center',
                                zIndex: 1000,
                                pointerEvents: 'auto'
                            }}>
                                <div style={{
                                    background: '#fff',
                                    padding: 30,
                                    borderRadius: 8,
                                    maxWidth: 600,
                                    width: '90%',
                                    maxHeight: '80vh',
                                    overflow: 'auto',
                                    zIndex: 1001,
                                    boxShadow: '0 4px 20px rgba(0,0,0,0.3)'
                                }}>
                                    <h2 style={{marginTop: 0}}>New Email</h2>

                                    {emailSent && (
                                        <div style={{
                                            padding: 10,
                                            marginBottom: 15,
                                            background: "#d4edda",
                                            color: "#155724",
                                            border: "1px solid #c3e6cb",
                                            borderRadius: 4
                                        }}>
                                            ✓ Email sent successfully!
                                        </div>
                                    )}

                                    <div style={{marginBottom: 15}}>
                                        <label
                                            style={{display: "block", marginBottom: 5, fontWeight: 600}}>From:</label>
                                        <input
                                            type="text"
                                            value={selectedUser}
                                            disabled
                                            style={{
                                                width: "100%",
                                                padding: 10,
                                                fontSize: 14,
                                                border: "1px solid #ccc",
                                                borderRadius: 4,
                                                background: "#f5f5f5"
                                            }}
                                        />
                                    </div>

                                    <div style={{marginBottom: 15}}>
                                        <label style={{display: "block", marginBottom: 5, fontWeight: 600}}>To:</label>
                                        <div style={{
                                            display: "flex",
                                            flexWrap: "wrap",
                                            gap: 6,
                                            alignItems: "center",
                                            padding: 8,
                                            border: "1px solid #ccc",
                                            borderRadius: 4,
                                            minHeight: 40
                                        }}>
                                            {newEmailRecipients.map((r, idx) => {
                                                const alias = teams.find(t => t.email === r)?.alias;
                                                return (
                                                    <span key={idx} style={{
                                                        display: "inline-flex",
                                                        alignItems: "center",
                                                        gap: 4,
                                                        background: "#e3f2fd",
                                                        padding: "4px 8px",
                                                        borderRadius: 12,
                                                        fontSize: 13
                                                    }}>
                                                        {alias ? `${alias} (${r})` : r}
                                                        <button
                                                            onClick={() => setNewEmailRecipients(prev => prev.filter((_, i) => i !== idx))}
                                                            style={{
                                                                background: "none",
                                                                border: "none",
                                                                cursor: "pointer",
                                                                color: "#999",
                                                                fontSize: 14,
                                                                padding: 0,
                                                                lineHeight: 1
                                                            }}
                                                        >✕</button>
                                                    </span>
                                                );
                                            })}
                                            <button
                                                onClick={() => setShowRecipientPicker(true)}
                                                style={{
                                                    background: "none",
                                                    border: "1px dashed #aaa",
                                                    borderRadius: 12,
                                                    padding: "4px 12px",
                                                    cursor: "pointer",
                                                    fontSize: 13,
                                                    color: "#1a73e8"
                                                }}
                                            >
                                                + Add Recipient
                                            </button>
                                        </div>

                                        {/* Recipient picker modal */}
                                        {showRecipientPicker && (
                                            <div style={{
                                                position: "fixed",
                                                inset: 0,
                                                background: "rgba(0,0,0,0.4)",
                                                display: "flex",
                                                alignItems: "center",
                                                justifyContent: "center",
                                                zIndex: 2000
                                            }}>
                                                <div style={{
                                                    background: "#fff",
                                                    padding: 20,
                                                    borderRadius: 8,
                                                    maxWidth: 400,
                                                    width: "90%",
                                                    maxHeight: "50vh",
                                                    display: "flex",
                                                    flexDirection: "column",
                                                    boxShadow: "0 4px 20px rgba(0,0,0,0.3)"
                                                }}>
                                                    <h3 style={{ margin: "0 0 12px 0" }}>Select Recipients</h3>
                                                    <div style={{ overflowY: "auto", flex: 1 }}>
                                                        {(() => {
                                                            const currentUser = users.find(u => u.email === selectedUser);
                                                            const managedTeams = currentUser?.teams || [];
                                                            if (managedTeams.length === 0) {
                                                                return <p style={{ color: "#999" }}>No managed teams available</p>;
                                                            }
                                                            return managedTeams.map(t => {
                                                                const teamData = teams.find(tm => tm.email === t.email);
                                                                const alias = teamData?.alias;
                                                                const isSelected = newEmailRecipients.includes(t.email);
                                                                return (
                                                                    <label
                                                                        key={t.email}
                                                                        style={{
                                                                            display: "flex",
                                                                            alignItems: "center",
                                                                            gap: 10,
                                                                            padding: "10px 8px",
                                                                            borderBottom: "1px solid #eee",
                                                                            cursor: "pointer",
                                                                            background: isSelected ? "#e3f2fd" : "transparent",
                                                                            borderRadius: 4
                                                                        }}
                                                                    >
                                                                        <input
                                                                            type="checkbox"
                                                                            checked={isSelected}
                                                                            onChange={() => {
                                                                                if (isSelected) {
                                                                                    setNewEmailRecipients(prev => prev.filter(r => r !== t.email));
                                                                                } else {
                                                                                    setNewEmailRecipients(prev => [...prev, t.email]);
                                                                                }
                                                                            }}
                                                                            style={{ width: 18, height: 18 }}
                                                                        />
                                                                        <span style={{ fontSize: 14 }}>
                                                                            {alias ? `${alias} (${t.email})` : t.email}
                                                                        </span>
                                                                    </label>
                                                                );
                                                            });
                                                        })()}
                                                    </div>
                                                    <button
                                                        onClick={() => setShowRecipientPicker(false)}
                                                        style={{
                                                            marginTop: 12,
                                                            padding: "8px 16px",
                                                            background: "#007bff",
                                                            color: "white",
                                                            border: "none",
                                                            borderRadius: 4,
                                                            cursor: "pointer",
                                                            fontSize: 14
                                                        }}
                                                    >
                                                        Done
                                                    </button>
                                                </div>
                                            </div>
                                        )}
                                    </div>

                                    <div style={{marginBottom: 15}}>
                                        <label style={{
                                            display: "block",
                                            marginBottom: 5,
                                            fontWeight: 600
                                        }}>Subject:</label>
                                        <input
                                            type="text"
                                            value={newEmailSubject}
                                            onChange={(e) => setNewEmailSubject(e.target.value)}
                                            placeholder="Email subject"
                                            style={{
                                                width: "100%",
                                                padding: 10,
                                                fontSize: 14,
                                                border: "1px solid #ccc",
                                                borderRadius: 4
                                            }}
                                        />
                                    </div>

                                    <div style={{marginBottom: 15}}>
                                        <label
                                            style={{display: "block", marginBottom: 5, fontWeight: 600}}>Body:</label>
                                        <textarea
                                            value={newEmailBody}
                                            onChange={(e) => setNewEmailBody(e.target.value)}
                                            placeholder="Type your message here..."
                                            rows={12}
                                            style={{
                                                width: "100%",
                                                padding: 10,
                                                fontSize: 14,
                                                border: "1px solid #ccc",
                                                borderRadius: 4,
                                                fontFamily: "inherit",
                                                resize: "vertical"
                                            }}
                                        />
                                    </div>

                                    <div style={{marginBottom: 15}}>
                                        <label style={{
                                            display: "inline-flex",
                                            alignItems: "center",
                                            gap: 6,
                                            cursor: "pointer",
                                            fontWeight: 600,
                                            fontSize: 14,
                                            color: "#1a73e8",
                                            padding: "8px 16px",
                                            border: "1px solid #1a73e8",
                                            borderRadius: 4,
                                            background: "white"
                                        }}>
                                            📎 Attach files
                                            <input
                                                type="file"
                                                multiple
                                                onChange={(e) => {
                                                    if (e.target.files) {
                                                        setNewEmailAttachments(prev => [...prev, ...Array.from(e.target.files!)]);
                                                    }
                                                }}
                                                style={{display: "none"}}
                                            />
                                        </label>
                                        {newEmailAttachments.length > 0 && (
                                            <div style={{marginTop: 6}}>
                                                {newEmailAttachments.map((file, idx) => (
                                                    <div key={idx} style={{
                                                        display: "flex",
                                                        alignItems: "center",
                                                        gap: 8,
                                                        padding: "4px 8px",
                                                        marginBottom: 4,
                                                        background: "#f0f0f0",
                                                        borderRadius: 4,
                                                        fontSize: 12
                                                    }}>
                                                <span style={{
                                                    flex: 1,
                                                    overflow: "hidden",
                                                    textOverflow: "ellipsis",
                                                    whiteSpace: "nowrap"
                                                }}>
                                                    {file.name} ({(file.size / 1024).toFixed(1)} KB)
                                                </span>
                                                        <button
                                                            onClick={() => setNewEmailAttachments(prev => prev.filter((_, i) => i !== idx))}
                                                            style={{
                                                                background: "none",
                                                                border: "none",
                                                                cursor: "pointer",
                                                                color: "#dc3545",
                                                                fontSize: 14,
                                                                padding: 0
                                                            }}
                                                        >✕
                                                        </button>
                                                    </div>
                                                ))}
                                            </div>
                                        )}
                                    </div>

                                    <div style={{fontSize: 12, color: "#666", marginBottom: 15}}>
                                        💡 Tip: Click on a template on the right to auto-fill subject and body
                                    </div>

                                    <div style={{display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center"}}>
                                        <button
                                            onClick={sendNewEmail}
                                            disabled={newEmailRecipients.length === 0 || !newEmailSubject.trim() || !newEmailBody.trim()}
                                            style={{
                                                padding: "10px 20px",
                                                background: (newEmailRecipients.length === 0 || !newEmailSubject.trim() || !newEmailBody.trim()) ? "#ccc" : "#28a745",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: (newEmailRecipients.length === 0 || !newEmailSubject.trim() || !newEmailBody.trim()) ? "not-allowed" : "pointer",
                                                fontSize: 14,
                                                fontWeight: 600
                                            }}
                                        >
                                            Send Email
                                        </button>
                                        <button
                                            onClick={() => setShowScheduleModal(!showScheduleModal)}
                                            disabled={isSchedulingEmail || newEmailRecipients.length === 0 || !newEmailSubject.trim() || !newEmailBody.trim()}
                                            style={{
                                                padding: "10px 20px",
                                                background: (isSchedulingEmail || newEmailRecipients.length === 0 || !newEmailSubject.trim() || !newEmailBody.trim()) ? "#ccc" : "#f0ad4e",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: (isSchedulingEmail || newEmailRecipients.length === 0 || !newEmailSubject.trim() || !newEmailBody.trim()) ? "not-allowed" : "pointer",
                                                fontSize: 14,
                                                fontWeight: 600
                                            }}
                                        >
                                            🕐 Schedule Send
                                        </button>
                                        <button
                                            onClick={() => {
                                                setShowNewEmailModal(false);
                                                setNewEmailRecipients([]);
                                                setNewEmailSubject('');
                                                setNewEmailBody('');
                                                setNewEmailAttachments([]);
                                                setShowScheduleModal(false);
                                                setScheduleTime('');
                                            }}
                                            style={{
                                                padding: "10px 20px",
                                                background: "#6c757d",
                                                color: "white",
                                                border: "none",
                                                borderRadius: 4,
                                                cursor: "pointer",
                                                fontSize: 14
                                            }}
                                        >
                                            Cancel
                                        </button>
                                    </div>

                                    {showScheduleModal && (
                                        <div style={{
                                            marginTop: 15,
                                            padding: 15,
                                            border: "1px solid #f0ad4e",
                                            borderRadius: 6,
                                            background: "#fff9e6"
                                        }}>
                                            <label style={{
                                                display: "block",
                                                marginBottom: 8,
                                                fontWeight: 600,
                                                fontSize: 14
                                            }}>
                                                Schedule for today at:
                                            </label>
                                            <input
                                                type="time"
                                                value={scheduleTime}
                                                onChange={(e) => setScheduleTime(e.target.value)}
                                                style={{
                                                    padding: 8,
                                                    fontSize: 14,
                                                    border: "1px solid #ccc",
                                                    borderRadius: 4,
                                                    marginRight: 10
                                                }}
                                            />
                                            <button
                                                onClick={() => {
                                                    if (isSchedulingEmail) return;
                                                    if (!scheduleTime) {
                                                        alert('Please select a time');
                                                        return;
                                                    }
                                                    const [hours, minutes] = scheduleTime.split(':');
                                                    const sendDate = new Date();
                                                    sendDate.setHours(parseInt(hours), parseInt(minutes), 0, 0);
                                                    if (sendDate <= new Date()) {
                                                        alert('Scheduled time must be in the future');
                                                        return;
                                                    }
                                                    scheduleEmail(sendDate.toISOString());
                                                }}
                                                disabled={isSchedulingEmail || !scheduleTime}
                                                style={{
                                                    padding: "8px 16px",
                                                    background: (isSchedulingEmail || !scheduleTime) ? "#ccc" : "#f0ad4e",
                                                    color: "white",
                                                    border: "none",
                                                    borderRadius: 4,
                                                    cursor: (isSchedulingEmail || !scheduleTime) ? "not-allowed" : "pointer",
                                                    fontSize: 14,
                                                    fontWeight: 600
                                                }}
                                            >
                                                {isSchedulingEmail ? 'Scheduling...' : 'Confirm Schedule'}
                                            </button>
                                        </div>
                                    )}
                                </div>
                            </div>
                        )}

                    </div>
                </div>

                {(attachmentPreviewLoading || attachmentPreview) && (
                    <div
                        onClick={clearAttachmentPreview}
                        style={{
                            position: 'fixed',
                            inset: 0,
                            background: 'rgba(0,0,0,0.6)',
                            display: 'flex',
                            alignItems: 'center',
                            justifyContent: 'center',
                            zIndex: 2500,
                            padding: 20
                        }}
                    >
                        <div
                            onClick={(e) => e.stopPropagation()}
                            style={{
                                width: 'min(1100px, 96vw)',
                                height: 'min(85vh, 900px)',
                                background: '#fff',
                                borderRadius: 8,
                                overflow: 'hidden',
                                boxShadow: '0 10px 30px rgba(0,0,0,0.35)',
                                display: 'flex',
                                flexDirection: 'column'
                            }}
                        >
                            <div style={{
                                display: 'flex',
                                justifyContent: 'space-between',
                                alignItems: 'center',
                                padding: '10px 14px',
                                borderBottom: '1px solid #e5e7eb',
                                background: '#f8f9fa'
                            }}>
                                <span style={{ fontSize: 13, fontWeight: 600 }}>
                                    {attachmentPreviewLoading ? 'Loading preview...' : `Preview: ${attachmentPreview?.name || ''}`}
                                </span>
                                <button
                                    onClick={clearAttachmentPreview}
                                    style={{
                                        border: 'none',
                                        background: 'transparent',
                                        cursor: 'pointer',
                                        color: '#666',
                                        fontSize: 20,
                                        lineHeight: 1
                                    }}
                                    title="Close preview"
                                >
                                    ×
                                </button>
                            </div>

                            <div style={{ flex: 1, overflow: 'auto', background: '#fff', padding: 8 }}>
                                {!attachmentPreviewLoading && attachmentPreview && (
                                    (() => {
                                        const previewMime = (attachmentPreview.mimeType || '').toLowerCase();
                                        const ext = getAttachmentExtension(attachmentPreview.name);

                                        if (attachmentPreview.previewKind === 'pptx-visual' && attachmentPreview.arrayBuffer) {
                                            return (
                                                <div style={{ padding: 10, display: 'flex', flexDirection: 'column', gap: 10, height: '100%' }}>
                                                    <div style={{
                                                        padding: '8px 10px',
                                                        fontSize: 12,
                                                        color: '#3f4d5a',
                                                        background: '#f5f7fa',
                                                        border: '1px solid #dce3eb',
                                                        borderRadius: 4
                                                    }}>
                                                        Visual PPTX rendering. If a slide looks wrong, use Download for exact fidelity.
                                                    </div>
                                                    <div style={{ flex: 1, minHeight: 500, border: '1px solid #e5e7eb', borderRadius: 6, overflow: 'auto' }}>
                                                        <PptxVisualPreview arrayBuffer={attachmentPreview.arrayBuffer} />
                                                    </div>
                                                </div>
                                            );
                                        }

                                        if (attachmentPreview.previewKind === 'docx-text') {
                                            return (
                                                <div style={{ padding: 10 }}>
                                                    <div style={{
                                                        padding: '8px 10px',
                                                        fontSize: 12,
                                                        color: '#3f4d5a',
                                                        background: '#f5f7fa',
                                                        border: '1px solid #dce3eb',
                                                        borderRadius: 4,
                                                        marginBottom: 10
                                                    }}>
                                                        Showing text preview extracted from DOCX.
                                                    </div>
                                                    <pre style={{
                                                        margin: 0,
                                                        whiteSpace: 'pre-wrap',
                                                        lineHeight: 1.45,
                                                        fontSize: 13,
                                                        fontFamily: 'inherit'
                                                    }}>
                                                        {attachmentPreview.parsedText || '(No readable text found)'}
                                                    </pre>
                                                </div>
                                            );
                                        }

                                        if (previewMime.startsWith('image/')) {
                                            return (
                                                <img
                                                    src={attachmentPreview.url}
                                                    alt={attachmentPreview.name}
                                                    style={{ maxWidth: '100%', maxHeight: '100%', display: 'block', margin: '0 auto' }}
                                                />
                                            );
                                        }

                                        if (isAudioMime(previewMime, attachmentPreview.name)) {
                                            return (
                                                <div style={{ maxWidth: 720, margin: '40px auto' }}>
                                                    <audio controls style={{ width: '100%' }} src={attachmentPreview.url}>
                                                        Your browser cannot play this audio preview.
                                                    </audio>
                                                </div>
                                            );
                                        }

                                        if (isVideoMime(previewMime, attachmentPreview.name)) {
                                            return (
                                                <div style={{ maxWidth: 960, margin: '20px auto' }}>
                                                    <video controls style={{ width: '100%', maxHeight: '70vh' }} src={attachmentPreview.url}>
                                                        Your browser cannot play this video preview.
                                                    </video>
                                                </div>
                                            );
                                        }

                                        if (isOfficeAttachment(previewMime, attachmentPreview.name)) {
                                            return (
                                                <div style={{ height: '100%', display: 'flex', flexDirection: 'column', gap: 8 }}>
                                                    <div style={{
                                                        padding: '8px 10px',
                                                        fontSize: 12,
                                                        color: '#3f4d5a',
                                                        background: '#f5f7fa',
                                                        border: '1px solid #dce3eb',
                                                        borderRadius: 4
                                                    }}>
                                                        Browser preview for {ext.toUpperCase() || 'Office'} files may be limited. If blank, use Download.
                                                    </div>
                                                    <iframe
                                                        src={attachmentPreview.url}
                                                        title={attachmentPreview.name}
                                                        style={{ width: '100%', height: '100%', minHeight: 500, border: 'none' }}
                                                    />
                                                </div>
                                            );
                                        }

                                        return (
                                            <iframe
                                                src={attachmentPreview.url}
                                                title={attachmentPreview.name}
                                                style={{ width: '100%', height: '100%', minHeight: 500, border: 'none' }}
                                            />
                                        );
                                    })()
                                )}
                            </div>
                        </div>
                    </div>
                )}

                {/* Create Template Modal */}
                {showCreateTemplateModal && (
                    <div style={{
                        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000
                    }}>
                        <div style={{
                            background: '#fff', padding: 24, borderRadius: 8,
                            maxWidth: 500, width: '90%', maxHeight: '80vh', overflow: 'auto',
                            boxShadow: '0 4px 20px rgba(0,0,0,0.3)'
                        }}>
                            <h3 style={{ marginTop: 0 }}>Create Template</h3>
                            <div style={{ marginBottom: 12 }}>
                                <label style={{ display: 'block', marginBottom: 4, fontWeight: 600, fontSize: 13 }}>Name</label>
                                <input type="text" value={templateName} onChange={e => setTemplateName(e.target.value)}
                                    placeholder="Template name"
                                    style={{ width: '100%', padding: 8, fontSize: 14, border: '1px solid #ccc', borderRadius: 4, boxSizing: 'border-box' }} />
                            </div>
                            <div style={{ marginBottom: 12 }}>
                                <label style={{ display: 'block', marginBottom: 4, fontWeight: 600, fontSize: 13 }}>Subject</label>
                                <input type="text" value={templateSubject} onChange={e => setTemplateSubject(e.target.value)}
                                    placeholder="Email subject"
                                    style={{ width: '100%', padding: 8, fontSize: 14, border: '1px solid #ccc', borderRadius: 4, boxSizing: 'border-box' }} />
                            </div>
                            <div style={{ marginBottom: 16 }}>
                                <label style={{ display: 'block', marginBottom: 4, fontWeight: 600, fontSize: 13 }}>Body</label>
                                <textarea value={templateBody} onChange={e => setTemplateBody(e.target.value)}
                                    placeholder="Email body"
                                    rows={8}
                                    style={{ width: '100%', padding: 8, fontSize: 14, border: '1px solid #ccc', borderRadius: 4, fontFamily: 'inherit', resize: 'vertical', boxSizing: 'border-box' }} />
                            </div>
                            <div style={{ display: 'flex', gap: 10 }}>
                                <button onClick={createTemplate}
                                    disabled={!templateName.trim() || !templateSubject.trim() || !templateBody.trim()}
                                    style={{
                                        padding: '8px 16px', background: (!templateName.trim() || !templateSubject.trim() || !templateBody.trim()) ? '#ccc' : '#28a745',
                                        color: 'white', border: 'none', borderRadius: 4,
                                        cursor: (!templateName.trim() || !templateSubject.trim() || !templateBody.trim()) ? 'not-allowed' : 'pointer', fontSize: 14
                                    }}>Create</button>
                                <button onClick={() => { setShowCreateTemplateModal(false); setTemplateName(''); setTemplateSubject(''); setTemplateBody(''); }}
                                    style={{ padding: '8px 16px', background: '#6c757d', color: 'white', border: 'none', borderRadius: 4, cursor: 'pointer', fontSize: 14 }}>
                                    Cancel
                                </button>
                            </div>
                        </div>
                    </div>
                )}

                {/* Edit Template Modal */}
                {showEditTemplateModal && editingTemplate && (
                    <div style={{
                        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000
                    }}>
                        <div style={{
                            background: '#fff', padding: 24, borderRadius: 8,
                            maxWidth: 500, width: '90%', maxHeight: '80vh', overflow: 'auto',
                            boxShadow: '0 4px 20px rgba(0,0,0,0.3)'
                        }}>
                            <h3 style={{ marginTop: 0 }}>Edit Template</h3>
                            <div style={{ marginBottom: 12 }}>
                                <label style={{ display: 'block', marginBottom: 4, fontWeight: 600, fontSize: 13 }}>Name</label>
                                <input type="text" value={templateName} onChange={e => setTemplateName(e.target.value)}
                                    style={{ width: '100%', padding: 8, fontSize: 14, border: '1px solid #ccc', borderRadius: 4, boxSizing: 'border-box' }} />
                            </div>
                            <div style={{ marginBottom: 12 }}>
                                <label style={{ display: 'block', marginBottom: 4, fontWeight: 600, fontSize: 13 }}>Subject</label>
                                <input type="text" value={templateSubject} onChange={e => setTemplateSubject(e.target.value)}
                                    style={{ width: '100%', padding: 8, fontSize: 14, border: '1px solid #ccc', borderRadius: 4, boxSizing: 'border-box' }} />
                            </div>
                            <div style={{ marginBottom: 16 }}>
                                <label style={{ display: 'block', marginBottom: 4, fontWeight: 600, fontSize: 13 }}>Body</label>
                                <textarea value={templateBody} onChange={e => setTemplateBody(e.target.value)}
                                    rows={8}
                                    style={{ width: '100%', padding: 8, fontSize: 14, border: '1px solid #ccc', borderRadius: 4, fontFamily: 'inherit', resize: 'vertical', boxSizing: 'border-box' }} />
                            </div>
                            <div style={{ display: 'flex', gap: 10 }}>
                                <button onClick={updateTemplate}
                                    disabled={!templateName.trim() || !templateSubject.trim() || !templateBody.trim()}
                                    style={{
                                        padding: '8px 16px', background: (!templateName.trim() || !templateSubject.trim() || !templateBody.trim()) ? '#ccc' : '#007bff',
                                        color: 'white', border: 'none', borderRadius: 4,
                                        cursor: (!templateName.trim() || !templateSubject.trim() || !templateBody.trim()) ? 'not-allowed' : 'pointer', fontSize: 14
                                    }}>Update</button>
                                <button onClick={() => { setShowEditTemplateModal(false); setEditingTemplate(null); setTemplateName(''); setTemplateSubject(''); setTemplateBody(''); }}
                                    style={{ padding: '8px 16px', background: '#6c757d', color: 'white', border: 'none', borderRadius: 4, cursor: 'pointer', fontSize: 14 }}>
                                    Cancel
                                </button>
                            </div>
                        </div>
                    </div>
                )}

                {/* Delete Template Confirmation Modal */}
                {showDeleteTemplateModal && deletingTemplate && (
                    <div style={{
                        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
                        display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 2000
                    }}>
                        <div style={{
                            background: '#fff', padding: 24, borderRadius: 8,
                            maxWidth: 400, width: '90%',
                            boxShadow: '0 4px 20px rgba(0,0,0,0.3)'
                        }}>
                            <h3 style={{ marginTop: 0 }}>Delete Template</h3>
                            <p>Are you sure you want to delete <strong>{deletingTemplate.name}</strong>?</p>
                            <div style={{ display: 'flex', gap: 10, justifyContent: 'flex-end', marginTop: 20 }}>
                                <button onClick={() => { setShowDeleteTemplateModal(false); setDeletingTemplate(null); }}
                                    style={{ padding: '8px 16px', background: '#6c757d', color: 'white', border: 'none', borderRadius: 4, cursor: 'pointer', fontSize: 14 }}>
                                    Cancel
                                </button>
                                <button onClick={deleteTemplate}
                                    style={{ padding: '8px 16px', background: '#dc3545', color: 'white', border: 'none', borderRadius: 4, cursor: 'pointer', fontSize: 14 }}>
                                    Delete
                                </button>
                            </div>
                        </div>
                    </div>
                )}

            </div>
        )
}

