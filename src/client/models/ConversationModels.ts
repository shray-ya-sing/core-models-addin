/**
 * Attachment type definition
 */
export interface Attachment {
  type: 'image' | 'pdf';
  name: string;
  content: string;
  mimeType: string;
}

/**
 * Message type definition
 */
export interface ChatMessage {
  role: 'user' | 'assistant' | 'system' | 'status';
  content: string;
  attachments?: Attachment[];
  isStreaming?: boolean;
  status?: string;
  stage?: string;
}

/**
 * Conversation session interface
 */
export interface ConversationSession {
  id: string;
  title: string;
  messages: ChatMessage[];
  lastUpdated: number;
  createdAt: number;
  workbookId?: string;
}

