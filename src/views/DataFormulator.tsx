// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import React, { useEffect, useRef, useState, useCallback, useMemo } from 'react';
import '../scss/App.scss';

import { useDispatch, useSelector } from "react-redux"; /* code change */
import { 
    DataFormulatorState,
    dfActions,
    dfSelectors,
} from '../app/dfSlice'

import _ from 'lodash';

import { Allotment, AllotmentHandle } from "allotment";
import "allotment/dist/style.css";

import {
    Typography,
    Box,
    Tooltip,
    Button,
    Divider,
    useTheme,
    alpha,
    CircularProgress,
    Backdrop,
    Link,
    Select,
    MenuItem,
    TextField,
} from '@mui/material';
import { borderColor, radius } from '../app/tokens';


import { VisualizationViewFC } from './VisualizationView';

import { DndProvider } from 'react-dnd'
import { HTML5Backend } from 'react-dnd-html5-backend'
import { toolName } from '../app/App';
import { DataThread } from './DataThread';

const dfLogo = '/rivus-logo-full.svg';
import exampleImageTable from "../assets/example-image-table.png";
import { ModelSelectionButton } from './ModelSelectionDialog';
import { UnifiedDataUploadDialog, UploadTabType, DataLoadMenu, ConnectorInstance } from './UnifiedDataUploadDialog';
import { ReportView } from './ReportView';
import { DataSourceSidebar } from './DataSourceSidebar';
import GitHubIcon from '@mui/icons-material/GitHub';
import { ExampleSession, exampleSessions, ExampleSessionCard, fetchExampleSessions } from './ExampleSessions';
import { useDataRefresh, useDerivedTableRefresh } from '../app/useDataRefresh';
import type { DictTable } from '../components/ComponentType';
import { useTranslation } from 'react-i18next';
import { fetchWithIdentity, getUrls, CONNECTOR_URLS } from '../app/utils';
import { apiRequest } from '../app/apiClient';
import { listWorkspaces, loadWorkspace, deleteWorkspace, exportWorkspace, importWorkspace, onWorkspaceListChanged, updateWorkspaceMeta } from '../app/workspaceService';
import type { WorkspaceSummary } from '../app/workspaceService';
import { AppDispatch } from '../app/store';
import { generateUUID } from '../app/identity';
import Card from '@mui/material/Card';
import CardContent from '@mui/material/CardContent';
import IconButton from '@mui/material/IconButton';
import DeleteOutlineIcon from '@mui/icons-material/DeleteOutline';
import DownloadIcon from '@mui/icons-material/Download';
import UploadFileIcon from '@mui/icons-material/UploadFile';
import EditOutlinedIcon from '@mui/icons-material/EditOutlined';
import ExpandMoreIcon from '@mui/icons-material/ExpandMore';
import Dialog from '@mui/material/Dialog';
import DialogTitle from '@mui/material/DialogTitle';
import DialogContent from '@mui/material/DialogContent';
import DialogActions from '@mui/material/DialogActions';

/** Generate a session ID like session_20260408_193052_a1b2 */
function generateSessionId(): string {
    const now = new Date();
    const date = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
    const time = `${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}`;
    const short = generateUUID().slice(0, 4);
    return `session_${date}_${time}_${short}`;
}

export const DataFormulatorFC = ({ }) => {

    const tables = useSelector((state: DataFormulatorState) => state.tables);
    const activeWorkspace = useSelector((state: DataFormulatorState) => state.activeWorkspace);
    const focusedId = useSelector((state: DataFormulatorState) => state.focusedId);
    const models = useSelector(dfSelectors.getAllModels);
    const selectedModelId = useSelector((state: DataFormulatorState) => state.selectedModelId);
    const viewMode = useSelector((state: DataFormulatorState) => state.viewMode);
    const serverConfig = useSelector((state: DataFormulatorState) => state.serverConfig);
    const identityKey = useSelector((state: DataFormulatorState) => `${state.identity.type}:${state.identity.id}`);
    const dataLoadingChatMessages = useSelector((state: DataFormulatorState) => state.dataLoadingChatMessages);
    const theme = useTheme();

    const dispatch = useDispatch<AppDispatch>();
    const { t } = useTranslation();

    // Auto-focus: when focusedId is undefined but tables exist, select the first table
    useEffect(() => {
        if (!focusedId && tables.length > 0) {
            dispatch(dfActions.setFocused({ type: 'table', tableId: tables[0].id }));
        }
    }, [focusedId, tables, dispatch]);

    // ── Connector instances (for landing page menu) ─────────────
    const [pageConnectors, setPageConnectors] = useState<ConnectorInstance[]>([]);
    const refreshPageConnectors = useCallback(() => {
        apiRequest<any>(CONNECTOR_URLS.LIST, { method: 'GET' })
            .then(({ data }) => setPageConnectors(data.connectors || []))
            .catch(() => { /* connector list is optional on landing page */ });
    }, []);
    const [connectorRefreshKey, setConnectorRefreshKey] = useState(0);
    const handleConnectorsChanged = useCallback(() => {
        setConnectorRefreshKey(k => k + 1);
        refreshPageConnectors();
    }, [refreshPageConnectors]);
    useEffect(() => {
        setPageConnectors([]);
        refreshPageConnectors();
    }, [refreshPageConnectors, identityKey]);

    // ── Demo sessions (loaded from manifest, fallback to hardcoded) ─────
    const [demoSessions, setDemoSessions] = useState<ExampleSession[]>(exampleSessions);
    useEffect(() => {
        fetchExampleSessions().then(sessions => {
            if (sessions.length > 0) setDemoSessions(sessions);
        });
    }, []);

    // ── Workspace list (shown on landing page) ────────────────────
    const [savedWorkspaces, setSavedWorkspaces] = useState<WorkspaceSummary[]>([]);
    const [confirmDeleteWs, setConfirmDeleteWs] = useState<string | null>(null);

    // Inline rename: which card's title is currently being edited, and
    // its draft text. Persisted via updateWorkspaceMeta on Enter / blur;
    // reverted on Escape.
    const [renamingWs, setRenamingWs] = useState<string | null>(null);
    const [renameDraft, setRenameDraft] = useState<string>('');

    // Sort key for the saved-workspaces grid. Default is creation time
    // so the user's chronological list of work doesn't shuffle every
    // time a workspace is touched.
    type WsSortKey = 'created_desc' | 'created_asc' | 'updated_desc' | 'name_asc';
    const [wsSort, setWsSort] = useState<WsSortKey>('created_desc');

    const fetchWorkspaces = useCallback(async () => {
        try {
            const sessions = await listWorkspaces();
            setSavedWorkspaces(sessions);
        } catch { /* workspace list is best-effort on landing page */ }
    }, []);

    useEffect(() => {
        if (!activeWorkspace || tables.length === 0) {
            fetchWorkspaces();
        }
    }, [activeWorkspace, tables.length, fetchWorkspaces]);

    useEffect(() => {
        return onWorkspaceListChanged(fetchWorkspaces);
    }, [fetchWorkspaces]);

    const handleOpenWorkspace = useCallback(async (name: string, metaDisplayName?: string) => {
        dispatch(dfActions.setSessionLoading({ loading: true, label: t('workspace.openingWorkspace') }));
        try {
            const result = await loadWorkspace(name);
            if (result && Object.keys(result.state).length > 0) {
                const displayName = metaDisplayName || result.displayName;
                dispatch(dfActions.loadState({ ...result.state, activeWorkspace: { id: name, displayName } }));
            } else {
                dispatch(dfActions.setActiveWorkspace({ id: name, displayName: metaDisplayName || 'Untitled Session' }));
            }
        } catch {
            dispatch(dfActions.setActiveWorkspace({ id: name, displayName: metaDisplayName || 'Untitled Session' }));
        }
        dispatch(dfActions.setSessionLoading({ loading: false }));
    }, [dispatch]);

    const handleDeleteWorkspace = useCallback(async (name: string) => {
        try {
            await deleteWorkspace(name);
            setSavedWorkspaces(prev => prev.filter(w => w.id !== name));
        } catch {
            dispatch(dfActions.addMessages({
                timestamp: Date.now(), type: 'error',
                component: 'workspace', value: t('workspace.deleteFailed'),
            }));
        }
        setConfirmDeleteWs(null);
    }, [dispatch]);

    const startRenameWorkspace = useCallback((id: string, currentName: string) => {
        setRenamingWs(id);
        setRenameDraft(currentName);
    }, []);

    const cancelRenameWorkspace = useCallback(() => {
        setRenamingWs(null);
        setRenameDraft('');
    }, []);

    const commitRenameWorkspace = useCallback(async () => {
        const id = renamingWs;
        if (!id) return;
        const next = renameDraft.trim();
        const current = savedWorkspaces.find(w => w.id === id);
        // Bail without writing if nothing changed or the new name is empty.
        if (!current || !next || next === current.display_name) {
            cancelRenameWorkspace();
            return;
        }
        // Optimistic update first so the UI reflects the change instantly;
        // the next list refresh (via onWorkspaceListChanged) will reconcile.
        setSavedWorkspaces(prev =>
            prev.map(w => (w.id === id ? { ...w, display_name: next } : w)),
        );
        cancelRenameWorkspace();
        try {
            await updateWorkspaceMeta(id, next);
        } catch {
            dispatch(dfActions.addMessages({
                timestamp: Date.now(), type: 'error',
                component: 'workspace', value: t('workspace.renameFailed'),
            }));
            // On failure, refetch so the UI returns to the server's truth.
            fetchWorkspaces();
        }
    }, [renamingWs, renameDraft, savedWorkspaces, cancelRenameWorkspace, dispatch, fetchWorkspaces]);

    const handleExportWorkspace = useCallback(async (id: string) => {
        try {
            const blob = await exportWorkspace(id);
            const ws = savedWorkspaces.find(w => w.id === id);
            const fileName = ws?.display_name || id;
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `${fileName}.zip`;
            a.click();
            URL.revokeObjectURL(a.href);
        } catch (e) {
            console.warn('Failed to export workspace:', e);
        }
    }, [savedWorkspaces]);

    const importRef = useRef<HTMLInputElement>(null);
    const handleImportWorkspace = useCallback(async (event: React.ChangeEvent<HTMLInputElement>) => {
        const file = event.target.files?.[0];
        if (!file) return;
        dispatch(dfActions.setSessionLoading({ loading: true, label: t('workspace.importingFile', { name: file.name }) }));
        try {
            const wsName = file.name.replace(/\.zip$/, '') || 'imported';
            const wsId = generateSessionId();
            const state = await importWorkspace(file, wsId, wsName);
            const restoredName = (state as any).activeWorkspace?.displayName || wsName;
            dispatch(dfActions.loadState({ ...state, activeWorkspace: { id: wsId, displayName: restoredName } }));
        } catch (e) {
            console.warn('Failed to import workspace:', e);
            dispatch(dfActions.addMessages({
                timestamp: Date.now(), type: 'error',
                component: 'workspace',
                value: t('workspace.importFailed'),
            }));
        }
        dispatch(dfActions.setSessionLoading({ loading: false }));
        if (importRef.current) importRef.current.value = '';
    }, [dispatch, t]);

    // Sorted view of saved workspaces. We don't mutate the underlying
    // list (the backend's response is the source of truth); we just
    // produce a re-ordered copy for rendering.
    const sortedSavedWorkspaces = useMemo(() => {
        const cmpDate = (a: string | null | undefined, b: string | null | undefined): number => {
            // Missing timestamps sort last regardless of direction so
            // legacy entries don't dominate either end of the list.
            if (!a && !b) return 0;
            if (!a) return 1;
            if (!b) return -1;
            return a.localeCompare(b);
        };
        const copy = [...savedWorkspaces];
        switch (wsSort) {
            case 'created_desc':
                return copy.sort((a, b) => cmpDate(b.created_at, a.created_at));
            case 'created_asc':
                return copy.sort((a, b) => cmpDate(a.created_at, b.created_at));
            case 'updated_desc':
                return copy.sort((a, b) => cmpDate(b.saved_at, a.saved_at));
            case 'name_asc':
                return copy.sort((a, b) =>
                    (a.display_name || '').localeCompare(b.display_name || ''),
                );
            default:
                return copy;
        }
    }, [savedWorkspaces, wsSort]);
    
    // Set up automatic refresh of derived tables when source data changes
    useDerivedTableRefresh();

    // State for unified data upload dialog
    const [uploadDialogOpen, setUploadDialogOpen] = useState(false);
    const [uploadDialogInitialTab, setUploadDialogInitialTab] = useState<UploadTabType>('menu');
    const [uploadDialogInitialChatPrompt, setUploadDialogInitialChatPrompt] = useState<string | undefined>(undefined);
    const [uploadDialogInitialChatImages, setUploadDialogInitialChatImages] = useState<string[] | undefined>(undefined);

    // Loading state for sessions (from Redux, shared with App.tsx)
    const sessionLoading = useSelector((state: DataFormulatorState) => state.sessionLoading);
    const sessionLoadingLabel = useSelector((state: DataFormulatorState) => state.sessionLoadingLabel);

    const openUploadDialog = (tab: UploadTabType, initialChatPrompt?: string, initialChatImages?: string[]) => {
        // If no workspace is active, generate an ID (backend creates folder lazily on first data op)
        if (!activeWorkspace) {
            dispatch(dfActions.setActiveWorkspace({ id: generateSessionId(), displayName: 'Untitled Session' }));
        }
        setUploadDialogInitialTab(tab);
        setUploadDialogInitialChatPrompt(initialChatPrompt);
        setUploadDialogInitialChatImages(initialChatImages);
        setUploadDialogOpen(true);
    };

    // Honor cross-component requests to hand off to the Data Loading
    // chat seeded with a prompt (e.g. Data Agent's `delegate` card with
    // target='data_loading'). Hand-offs targeting other agents (e.g.
    // `report_gen`) are consumed elsewhere — we only clear our own.
    const agentHandoffRequest = useSelector((state: DataFormulatorState) => state.agentHandoffRequest);
    useEffect(() => {
        if (agentHandoffRequest && agentHandoffRequest.target === 'data_loading') {
            openUploadDialog('extract', agentHandoffRequest.prompt, agentHandoffRequest.images);
            dispatch(dfActions.clearAgentHandoffRequest());
        }
        // openUploadDialog is stable enough for this purpose; we only react
        // to changes in the handoff request itself.
        // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [agentHandoffRequest]);

    const handleLoadExampleSession = async (session: ExampleSession) => {
        dispatch(dfActions.setSessionLoading({ loading: true, label: t('messages.loadingExample', { title: session.title }) }));

        dispatch(dfActions.addMessages({
            timestamp: Date.now(),
            type: 'info',
            component: 'data formulator',
            value: t('messages.loadingExample', { title: session.title }),
        }));

        try {
            // Fetch the workspace zip
            const res = await fetch(session.workspace);
            if (!res.ok) throw new Error(`Failed to fetch ${session.workspace}`);
            const blob = await res.blob();
            const file = new File([blob], `${session.id}.zip`, { type: 'application/zip' });

            // Import via the standard workspace import flow (parquet + state)
            const wsId = generateSessionId();
            // Set workspace ID first so fetchWithIdentity sends X-Workspace-Id header
            dispatch(dfActions.setActiveWorkspace({ id: wsId, displayName: session.title }));
            const state = await importWorkspace(file, wsId, session.title);
            dispatch(dfActions.loadState({ ...state, activeWorkspace: { id: wsId, displayName: session.title } }));

            dispatch(dfActions.addMessages({
                timestamp: Date.now(),
                type: 'success',
                component: 'data formulator',
                value: t('messages.loadSuccess', { title: session.title }),
            }));
        } catch (error: any) {
            console.error('Error loading session:', error);
            dispatch(dfActions.addMessages({
                timestamp: Date.now(),
                type: 'error',
                component: 'data formulator',
                value: t('messages.loadFailed', { title: session.title, error: error.message }),
            }));
        } finally {
            dispatch(dfActions.setSessionLoading({ loading: false }));
        }
    };

    useEffect(() => {
        document.title = toolName;
        
        // Preload imported images (public images are preloaded in index.html)
        const imagesToPreload = [
            { src: dfLogo, type: 'image/svg+xml' },
            { src: exampleImageTable, type: 'image/png' },
        ];
        
        const preloadLinks: HTMLLinkElement[] = [];
        imagesToPreload.forEach(({ src, type }) => {
            // Use link preload for better priority
            const link = document.createElement('link');
            link.rel = 'preload';
            link.as = 'image';
            link.href = src;
            link.type = type;
            document.head.appendChild(link);
            preloadLinks.push(link);
        });
        
        // Cleanup function to remove preload links when component unmounts
        return () => {
            preloadLinks.forEach(link => {
                if (link.parentNode) {
                    link.parentNode.removeChild(link);
                }
            });
        };
    }, []);

    useEffect(() => {
        // Auto-select the first available model when none is selected.
        // No connectivity check on load — errors surface on first use,
        // and the user can manually test via the model selection dialog.
        if (selectedModelId === undefined && models.length > 0) {
            dispatch(dfActions.selectModel(models[0].id));
        }
    }, [dispatch, models, selectedModelId]);

    const visPaneMain = (
        <Box sx={{ width: "100%", height: "100%", overflow: "hidden", display: "flex", flexDirection: "row" }}>
            <VisualizationViewFC />
        </Box>);

    const visPane = visPaneMain;

    let borderBoxStyle = {
        border: `1px solid ${borderColor.view}`, 
        borderRadius: radius.pill, 
        //boxShadow: '0 0 5px rgba(0,0,0,0.1)',
    }

    // Discrete column snapping for DataThread
    const CARD_WIDTH = 220;
    const CARD_GAP = 12;
    const COLUMN_WIDTH = CARD_WIDTH + CARD_GAP;
    const PANE_PADDING = 48;
    const columnSize = (n: number) => n * COLUMN_WIDTH + PANE_PADDING;
    const allotmentRef = useRef<AllotmentHandle>(null);
    const containerRef = useRef<HTMLDivElement>(null);

    const snapToColumns = useCallback((sizes: number[]) => {
        if (!allotmentRef.current || sizes.length < 2) return;
        const raw = sizes[0];
        // Find nearest discrete column count (1-3)
        let bestCols = 1;
        let bestDist = Infinity;
        for (let n = 1; n <= 3; n++) {
            const dist = Math.abs(raw - columnSize(n));
            if (dist < bestDist) {
                bestDist = dist;
                bestCols = n;
            }
        }
        const snapped = columnSize(bestCols);
        if (Math.abs(raw - snapped) > 2) {
            const totalWidth = sizes.reduce((a, b) => a + b, 0);
            allotmentRef.current.resize([snapped, totalWidth - snapped]);
        }
    }, []);

    // Compute thread count to decide preferred pane width:
    // A "thread" is a leaf table's derivation chain displayed as a column.
    // Must match the chain-splitting logic in DataThread (MAX_CHAIN_TABLES).
    const threadCount = useMemo(() => {
        // A table is a "leaf" if no other non-anchored table derives from it
        const hasNonAnchoredChild = new Set<string>();
        tables.forEach(t => {
            if (t.derive && !t.anchored) {
                hasNonAnchoredChild.add(t.derive.trigger.tableId);
            }
        });
        const leafTables = tables.filter(t => !hasNonAnchoredChild.has(t.id));
        // Threads = leaf tables with derivation chains + 1 group for hanging (source) tables
        const threaded = leafTables.filter(t => t.derive);
        const hanging = leafTables.filter(t => !t.derive);
        let count = threaded.length + (hanging.length > 0 ? 1 : 0);

        // Account for chain-splitting: long chains are broken into sub-threads
        // (mirrors MAX_CHAIN_TABLES logic in DataThread)
        const MAX_CHAIN_TABLES = 5;
        const tableById = new Map(tables.map(t => [t.id, t]));
        const getChainLength = (t: DictTable): number => {
            let len = 1;
            let cur = t;
            while (cur.derive && !cur.anchored) {
                len++;
                const parent = tableById.get(cur.derive.trigger.tableId);
                if (!parent) break;
                cur = parent;
            }
            return len;
        };
        const claimedForCount = new Set<string>();
        for (const lt of threaded) {
            // Walk chain
            const chainIds: string[] = [lt.id];
            let cur = lt;
            while (cur.derive && !cur.anchored) {
                const pid = cur.derive.trigger.tableId;
                chainIds.push(pid);
                const parent = tableById.get(pid);
                if (!parent) break;
                cur = parent;
            }
            const ownedIds = chainIds.filter(id => !claimedForCount.has(id));
            if (ownedIds.length > MAX_CHAIN_TABLES) {
                // Each extra split adds one more thread entry
                const extraSplits = Math.floor((ownedIds.length - 1) / MAX_CHAIN_TABLES);
                count += extraSplits;
            }
            chainIds.forEach(id => claimedForCount.add(id));
        }

        return count;
    }, [tables]);
    const preferredColumns = threadCount <= 1 ? 1 : 2;

    // Track previous thread count to auto-resize intelligently
    const prevThreadCountRef = useRef(threadCount);
    useEffect(() => {
        const prev = prevThreadCountRef.current;
        prevThreadCountRef.current = threadCount;
        if (!allotmentRef.current || !containerRef.current) return;
        // When there are no tables the first Allotment.Pane is unmounted,
        // so the Allotment only has one child – calling resize with two
        // sizes would crash (accessing .minimumSize on an undefined pane).
        if (tables.length === 0) return;
        const totalWidth = containerRef.current.clientWidth;
        if (totalWidth <= 0) return;

        let newSize: number | null = null;
        if (prev <= 1 && threadCount > 1) {
            // Case 1: was 1 thread, now 2+ → expand to 2 columns
            newSize = columnSize(2);
        } else if (prev > 1 && threadCount <= 1) {
            // Case 2: was 2+ threads, now 1 → shrink to 1 column
            newSize = columnSize(1);
        }
        // Case 3: was 2+ threads and still 2+ → don't change (respect user's manual setting)

        if (newSize !== null) {
            // Defer resize to the next animation frame so the Allotment has
            // re-rendered its pane children before we call resize.
            const finalSize = newSize;
            const rafId = requestAnimationFrame(() => {
                try {
                    const w = containerRef.current?.clientWidth ?? totalWidth;
                    allotmentRef.current?.resize([finalSize, w - finalSize]);
                } catch {
                    // Allotment pane structure may not yet match; ignore.
                }
            });
            return () => cancelAnimationFrame(rafId);
        }
    }, [threadCount, tables.length]);

    const fixedSplitPane = ( 
        <Box sx={{display: 'flex', flexDirection: 'row', height: '100%'}}>
            <DataSourceSidebar
                onOpenUploadDialog={(tab) => openUploadDialog((tab ?? 'add-connection') as UploadTabType)}
                connectorRefreshKey={connectorRefreshKey}
            />
            <Box ref={containerRef} className="outer-allotment" sx={{
                    margin: '4px 8px 8px 8px', backgroundColor: 'white',
                    display: 'flex', height: 'calc(100% - 12px)', flex: 1, minWidth: 0, flexDirection: 'column',
                    overflow: 'hidden',
                    position: 'relative'}}>
                <Allotment ref={allotmentRef} onDragEnd={snapToColumns} proportionalLayout={false}>
                    {tables.length > 0 ? (
                        <Allotment.Pane minSize={columnSize(1)} preferredSize={columnSize(preferredColumns)} maxSize={columnSize(3)} snap={false}>
                            <DataThread sx={{
                                display: 'flex', 
                                flexDirection: 'column',
                                overflow: 'hidden',
                                alignContent: 'flex-start',
                                height: '100%',
                            }}/>
                        </Allotment.Pane>
                    ) : null}
                    <Allotment.Pane minSize={300}>
                        <Box sx={{ ...borderBoxStyle, height: '100%', overflow: 'hidden', display: 'flex', flexDirection: 'column', boxSizing: 'border-box' }}>
                            {viewMode === 'editor' ? (
                                visPane
                            ) : (
                                <ReportView />
                            )}
                        </Box>
                    </Allotment.Pane>
                </Allotment>
            </Box>
        </Box>
    );

    let footer = <></>;

    let dataUploadRequestBox = <Box sx={{
            margin: '4px 4px 4px 8px', 
            background: `
                linear-gradient(90deg, ${alpha(theme.palette.text.secondary, 0.01)} 1px, transparent 1px),
                linear-gradient(0deg, ${alpha(theme.palette.text.secondary, 0.01)} 1px, transparent 1px)
            `,
            backgroundSize: '16px 16px',
            flex: 1, minWidth: 0, overflow: 'auto', display: 'flex', flexDirection: 'column', height: '100%',
        }}>
        <Box sx={{margin:'auto', pb: '5%', display: "flex", flexDirection: "column", textAlign: "center", maxWidth: 1024, width: '100%', px: 2, boxSizing: 'border-box' }}>
            <Box sx={{display: 'flex', mx: 'auto'}}>
                <Typography fontSize={84} sx={{ml: 2, letterSpacing: '0.05em'}}>{toolName}</Typography> 
            </Box>
            <Typography sx={{ 
                fontSize: 24, color: theme.palette.text.secondary, 
                textAlign: 'center', mb: 2}}>
                {t('landing.tagline')}
            </Typography>

            <Box sx={{mt: 4}}>
                <DataLoadMenu 
                    onSelectTab={(tab) => openUploadDialog(tab)}
                    onSelectConnector={(conn) => {
                        // Already-authed connector → open the data-source
                        // sidebar focused on it. Otherwise open the upload
                        // dialog at the connector's auth/connect tab.
                        if (conn.connected || conn.sso_auto_connect) {
                            dispatch(dfActions.focusConnector(conn.id));
                        } else {
                            openUploadDialog(`connector:${conn.id}` as UploadTabType);
                        }
                    }}
                    onStartChat={(prompt, images) => openUploadDialog('extract', prompt, images)}
                    hasPriorConversation={dataLoadingChatMessages.length > 0}
                    onResumeChat={() => openUploadDialog('extract')}
                    serverConfig={serverConfig}
                    connectors={pageConnectors}
                />
            </Box>


            {/* ── Delete workspace confirmation ────────────────────── */}
            <Dialog open={confirmDeleteWs !== null} onClose={() => setConfirmDeleteWs(null)}>
                <DialogTitle>{t('workspace.deleteTitle')}</DialogTitle>
                <DialogContent>
                    <Typography dangerouslySetInnerHTML={{
                        __html: t('workspace.deleteConfirm', {
                            name: savedWorkspaces.find(w => w.id === confirmDeleteWs)?.display_name || confirmDeleteWs,
                            id: confirmDeleteWs,
                            interpolation: { escapeValue: false },
                        }),
                    }} />
                </DialogContent>
                <DialogActions>
                    <Button onClick={() => setConfirmDeleteWs(null)}>{t('workspace.cancel')}</Button>
                    <Button color="error" onClick={() => confirmDeleteWs && handleDeleteWorkspace(confirmDeleteWs)}>
                        {t('workspace.delete')}
                    </Button>
                </DialogActions>
            </Dialog>
        </Box>
        {footer}
    </Box>;
    
    return (
        <Box sx={{ display: 'block', width: "100%", height: '100%', position: 'relative' }}>
            <DndProvider backend={HTML5Backend}>
                {tables.length > 0 ? fixedSplitPane : (
                    <Box sx={{ display: 'flex', flexDirection: 'row', height: '100%' }}>
                        <DataSourceSidebar
                            onOpenUploadDialog={(tab) => openUploadDialog((tab ?? 'add-connection') as UploadTabType)}
                            connectorRefreshKey={connectorRefreshKey}
                        />
                        {dataUploadRequestBox}
                    </Box>
                )}
                <UnifiedDataUploadDialog 
                    open={uploadDialogOpen}
                    onClose={() => {
                        setUploadDialogOpen(false);
                        // Clear one-shot seed values so the next dialog
                        // open (e.g. via the upload button) doesn't
                        // re-fire the agent with a stale prompt/image.
                        setUploadDialogInitialChatPrompt(undefined);
                        setUploadDialogInitialChatImages(undefined);
                        refreshPageConnectors();
                    }}
                    initialTab={uploadDialogInitialTab}
                    initialChatPrompt={uploadDialogInitialChatPrompt}
                    initialChatImages={uploadDialogInitialChatImages}
                    onConnectorsChanged={handleConnectorsChanged}
                />
                {/* Loading overlay for session loading */}
                <Backdrop
                    open={sessionLoading}
                    sx={{
                        position: 'absolute',
                        zIndex: 999,
                        backgroundColor: alpha(theme.palette.background.default, 0.85),
                        backdropFilter: 'blur(4px)',
                        display: 'flex',
                        flexDirection: 'column',
                        gap: 2,
                    }}
                >
                    <CircularProgress size={40} />
                    <Typography variant="body1" color="text.secondary">
                        {sessionLoadingLabel || t('session.loadingSessions')}
                    </Typography>
                    <Button
                        variant="text"
                        size="small"
                        onClick={() => dispatch(dfActions.setSessionLoading({ loading: false }))}
                        sx={{ mt: 1, textTransform: 'none', color: 'text.secondary' }}
                    >
                        {t('app.cancel')}
                    </Button>
                </Backdrop>
                {selectedModelId == undefined && (
                    <Box sx={{
                        position: 'absolute',
                        top: 0,
                        left: 0,
                        right: 0,
                        bottom: 0,
                        backgroundColor: alpha(theme.palette.background.default, 0.85),
                        backdropFilter: 'blur(4px)',
                        display: 'flex',
                        flexDirection: 'column',
                        zIndex: 1000,
                    }}>
                        <Box sx={{margin:'auto', pb: '5%', display: "flex", flexDirection: "column", textAlign: "center"}}>
                            <Box component="img" sx={{  width: 196, margin: "auto" }} alt="Rivus logo" src={dfLogo} fetchPriority="high" />
                            <Typography variant="h3" sx={{marginTop: "20px", fontWeight: 200, letterSpacing: '0.05em'}}>
                                {toolName}
                            </Typography>
                            <Typography  variant="h4" sx={{mt: 3, fontSize: 28, letterSpacing: '0.02em'}}>
                                {t('landing.firstSelectModelPrefix')} <ModelSelectionButton />
                            </Typography>

                        </Box>
                        {footer}
                    </Box>
                )}
            </DndProvider>
        </Box>);
}