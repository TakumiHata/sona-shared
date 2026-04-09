import type { AgendaItem, FlatAgendaWithDepth } from '../types';
import { stripMeetingSummaryTags } from '../tags';

const AGENDA_TAG_RE = /<agenda[^>]*>([\s\S]*?)<\/agenda>/g;

/** テキスト内に <agenda> タグが含まれているか */
export function hasAgendaTags(text: string): boolean {
    return new RegExp(AGENDA_TAG_RE.source).test(text);
}

/**
 * Parses a Markdown string and returns a list of AgendaItems.
 *
 * - <agenda> タグがある場合: タグ内の見出しのみ議題として認識
 * - <agenda> タグがない場合: 従来通り全ての見出しを議題として認識（後方互換）
 */
export const parseAgendaMarkdown = (text: string): AgendaItem[] => {
    // <meeting-summary> はインポート時に除去（出力時のみ生成）
    const cleaned = stripMeetingSummaryTags(text);
    if (hasAgendaTags(cleaned)) {
        return parseWithAgendaTags(cleaned);
    }
    return parseLegacy(cleaned);
};

function parseWithAgendaTags(text: string): AgendaItem[] {
    const agendaItems: AgendaItem[] = [];
    let currentRoot: AgendaItem | null = null;
    let currentLevel2: AgendaItem | null = null;
    let currentItem: AgendaItem | null = null;

    const agendaRanges: { start: number; end: number }[] = [];
    const re = new RegExp(AGENDA_TAG_RE.source, 'g');
    let match;
    while ((match = re.exec(text)) !== null) {
        const tagStart = match.index;
        const contentStart = tagStart + match[0].length - match[1].length - '</agenda>'.length;
        const contentEnd = contentStart + match[1].length;
        agendaRanges.push({ start: contentStart, end: contentEnd });
    }

    let charOffset = 0;
    const lines = text.split('\n');

    for (const line of lines) {
        const lineStart = charOffset;
        const lineEnd = charOffset + line.length;
        charOffset = lineEnd + 1;

        const trimmedLine = line.trim();

        if (trimmedLine.startsWith('<agenda') || trimmedLine === '</agenda>') {
            continue;
        }

        const inAgenda = agendaRanges.some(r => lineStart >= r.start && lineEnd <= r.end);

        if (inAgenda && trimmedLine.startsWith('# ')) {
            const title = trimmedLine.substring(2).trim();
            const newItem: AgendaItem = { id: crypto.randomUUID(), title, children: [] };
            agendaItems.push(newItem);
            currentRoot = newItem;
            currentLevel2 = null;
            currentItem = newItem;
        } else if (inAgenda && trimmedLine.startsWith('## ')) {
            const title = trimmedLine.substring(3).trim();
            const newItem: AgendaItem = { id: crypto.randomUUID(), title, children: [] };
            if (currentRoot) {
                currentRoot.children = currentRoot.children || [];
                currentRoot.children.push(newItem);
            } else {
                agendaItems.push(newItem);
                currentRoot = newItem;
            }
            currentLevel2 = newItem;
            currentItem = newItem;
        } else if (inAgenda && trimmedLine.startsWith('### ')) {
            const title = trimmedLine.substring(4).trim();
            const newItem: AgendaItem = { id: crypto.randomUUID(), title, children: [] };
            if (currentLevel2) {
                currentLevel2.children = currentLevel2.children || [];
                currentLevel2.children.push(newItem);
            } else if (currentRoot) {
                currentRoot.children = currentRoot.children || [];
                currentRoot.children.push(newItem);
            } else {
                agendaItems.push(newItem);
                currentRoot = newItem;
            }
            currentItem = newItem;
        } else if (currentItem) {
            if (currentItem.description != null) {
                currentItem.description += '\n' + line;
            } else if (trimmedLine.length > 0) {
                currentItem.description = line;
            }
        }
    }

    setOriginalDescriptions(agendaItems);
    return agendaItems;
}

function parseLegacy(text: string): AgendaItem[] {
    const lines = text.split('\n');
    const agendaItems: AgendaItem[] = [];
    let currentRoot: AgendaItem | null = null;
    let currentLevel2: AgendaItem | null = null;
    let currentItem: AgendaItem | null = null;

    lines.forEach(line => {
        const trimmedLine = line.trim();

        if (trimmedLine.startsWith('# ')) {
            const title = trimmedLine.substring(2).trim();
            const newItem: AgendaItem = { id: crypto.randomUUID(), title, children: [] };
            agendaItems.push(newItem);
            currentRoot = newItem;
            currentLevel2 = null;
            currentItem = newItem;
        } else if (trimmedLine.startsWith('## ')) {
            const title = trimmedLine.substring(3).trim();
            const newItem: AgendaItem = { id: crypto.randomUUID(), title, children: [] };
            if (currentRoot) {
                currentRoot.children = currentRoot.children || [];
                currentRoot.children.push(newItem);
            } else {
                agendaItems.push(newItem);
                currentRoot = newItem;
            }
            currentLevel2 = newItem;
            currentItem = newItem;
        } else if (trimmedLine.startsWith('### ')) {
            const title = trimmedLine.substring(4).trim();
            const newItem: AgendaItem = { id: crypto.randomUUID(), title, children: [] };
            if (currentLevel2) {
                currentLevel2.children = currentLevel2.children || [];
                currentLevel2.children.push(newItem);
            } else if (currentRoot) {
                currentRoot.children = currentRoot.children || [];
                currentRoot.children.push(newItem);
            } else {
                agendaItems.push(newItem);
                currentRoot = newItem;
            }
            currentItem = newItem;
        } else if (currentItem) {
            if (currentItem.description != null) {
                currentItem.description += '\n' + line;
            } else if (trimmedLine.length > 0) {
                currentItem.description = line;
            }
        }
    });

    setOriginalDescriptions(agendaItems);
    return agendaItems;
}

function setOriginalDescriptions(items: AgendaItem[]) {
    for (const item of items) {
        if (item.description) {
            item.description = item.description.replace(/\n+$/, '');
            item.originalDescription = item.description;
        }
        if (item.children) setOriginalDescriptions(item.children);
    }
}

/** recursively searches for an agenda item by ID. */
export const findAgendaById = (agendas: AgendaItem[], id: string): AgendaItem | undefined => {
    for (const agenda of agendas) {
        if (agenda.id === id) return agenda;
        if (agenda.children) {
            const found = findAgendaById(agenda.children, id);
            if (found) return found;
        }
    }
    return undefined;
};

/** recursively searches for an agenda item by title. */
export const findAgendaByTitle = (agendas: AgendaItem[], title: string): AgendaItem | undefined => {
    for (const agenda of agendas) {
        if (agenda.title === title) return agenda;
        if (agenda.children) {
            const found = findAgendaByTitle(agenda.children, title);
            if (found) return found;
        }
    }
    return undefined;
};

/** Recursively updates an agenda item in the tree by ID. */
export const updateAgendaInTree = (
    agendas: AgendaItem[],
    id: string,
    updater: (item: AgendaItem) => AgendaItem
): AgendaItem[] => {
    return agendas.map(agenda => {
        if (agenda.id === id) return updater(agenda);
        if (agenda.children) {
            return { ...agenda, children: updateAgendaInTree(agenda.children, id, updater) };
        }
        return agenda;
    });
};

/** Flattens the agenda tree into a single array. */
export const flattenAgendas = (agendas: AgendaItem[]): AgendaItem[] => {
    let result: AgendaItem[] = [];
    for (const agenda of agendas) {
        result.push(agenda);
        if (agenda.children) {
            result = result.concat(flattenAgendas(agenda.children));
        }
    }
    return result;
};

/** Flattens the agenda tree with depth information. */
export const flattenAgendasWithDepth = (agendas: AgendaItem[], depth = 0): FlatAgendaWithDepth[] => {
    let result: FlatAgendaWithDepth[] = [];
    for (const agenda of agendas) {
        result.push({ ...agenda, depth });
        if (agenda.children) {
            result = result.concat(flattenAgendasWithDepth(agenda.children, depth + 1));
        }
    }
    return result;
};

/** Adds a new agenda item to the tree. */
export const addAgendaItem = (
    agendas: AgendaItem[],
    newItem: AgendaItem,
    parentId?: string
): AgendaItem[] => {
    if (!parentId) return [...agendas, newItem];

    return agendas.map(agenda => {
        if (agenda.id === parentId) {
            return { ...agenda, children: agenda.children ? [...agenda.children, newItem] : [newItem] };
        }
        if (agenda.children) {
            return { ...agenda, children: addAgendaItem(agenda.children, newItem, parentId) };
        }
        return agenda;
    });
};

/** Recursively deletes an agenda item from the tree by ID. */
export const deleteAgendaItem = (agendas: AgendaItem[], id: string): AgendaItem[] => {
    return agendas.reduce((acc, agenda) => {
        if (agenda.id === id) return acc;
        if (agenda.children) {
            acc.push({ ...agenda, children: deleteAgendaItem(agenda.children, id) });
        } else {
            acc.push(agenda);
        }
        return acc;
    }, [] as AgendaItem[]);
};

/**
 * アジェンダツリーからMarkdownテキストを再構築する。
 * 文字起こし結果（refinedTranscript / rawTranscript）を含めて出力する。
 */
export const buildMarkdownFromAgendas = (
    agendas: AgendaItem[],
    useAgendaTags: boolean
): string => {
    const sections: string[] = [];

    const addAgenda = (items: AgendaItem[], depth: number) => {
        for (const item of items) {
            const prefix = '#'.repeat(Math.min(depth + 1, 4));
            if (useAgendaTags) {
                sections.push(`<agenda data-id="${item.id}">\n${prefix} ${item.title}\n</agenda>`);
            } else {
                sections.push(`${prefix} ${item.title}`);
            }

            if (item.summaryText) {
                sections.push('');
                sections.push(`<meeting-summary>\n\n${item.summaryText}\n\n</meeting-summary>`);
            }

            if (item.refinedTranscript) {
                if (item.originalDescription) {
                    sections.push('');
                    sections.push(item.originalDescription);
                }
                sections.push('');
                sections.push(`<details>\n<summary>文字起こし結果</summary>\n\n${item.refinedTranscript}\n\n</details>`);
            } else if (item.rawTranscript) {
                if (item.originalDescription) {
                    sections.push('');
                    sections.push(item.originalDescription);
                }
                sections.push('');
                sections.push(`<details>\n<summary>文字起こし結果</summary>\n\n> ${item.rawTranscript.replace(/\n/g, '\n> ')}\n\n</details>`);
            } else if (item.description) {
                sections.push('');
                sections.push(item.description);
            }
            sections.push('');

            if (item.children) {
                addAgenda(item.children, depth + 1);
            }
        }
    };

    addAgenda(agendas, 0);
    return sections.join('\n');
};
