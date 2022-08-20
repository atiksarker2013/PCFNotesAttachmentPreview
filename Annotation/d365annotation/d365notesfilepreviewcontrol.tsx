import * as React from 'react';
import {
    getTheme, loadTheme, mergeStyleSets, FontWeights,
    DefaultButton, PrimaryButton, IconButton, IIconProps, IButtonStyles,
    CommandBar, ICommandBarItemProps,
    DetailsList, DetailsListLayoutMode, IColumn, IDetailsListStyles, Selection,
    ILabelStyles, Label,
    MarqueeSelection,
    Modal,
    Panel,
    Spinner, SpinnerSize,
    Stack, IStackProps,
    Text,
    TextField
} from '@fluentui/react';
import * as sanitize from 'sanitize-html';
import { IInputs } from './generated/ManifestTypes';
import ThemeProvider from './ThemeProvider';
const FileViewer = require('react-file-viewer');


export interface Id365notesfilepreviewControlProps {
    context: ComponentFramework.Context<IInputs>;
    themeMainColor: string;
    entityId: string;
    entityName: string;
}

export interface Id365notesfilepreviewControlState {
    columns: IColumn[];
    gridItems: INote[];
    dataLoaded: boolean;
    filesUploading: boolean;
    selectedItemIds: string[];
    sidePanelState: SidePanelState;
    previewOpen: boolean;
    noteId: string;
    noteTitle: string;
    noteText: string;
    noteFile: ComponentFramework.FileObject | null;
    previewedItem: INote | null;
}

export interface INote {
    Id: string;
    title: string;
    description: string;
    fileName: string;
    file?: ComponentFramework.FileObject;
    owner: string;
    dateCreated: Date;
}

export interface IFilesUpload {
    files: File[];
    createdNotesRefs: ComponentFramework.LookupValue[];
    errors: string[];
}

export enum SidePanelState {
    closed = 0,
    create = 1,
    edit = 2
}

export class d365notesfilepreviewControl extends React.Component<Id365notesfilepreviewControlProps, Id365notesfilepreviewControlState> {
    private _props: Id365notesfilepreviewControlProps;
    private _allItems: INote[];
    private _theme: any;
    private _selection: Selection;
    private _filesUpload: IFilesUpload;
    private _interval: NodeJS.Timeout;


    constructor(props: Id365notesfilepreviewControlProps) {
        super(props);

        this._props = props;
        this._allItems = [];
        this._filesUpload = {
            files: [],
            createdNotesRefs: [],
            errors: []
        };

        this._theme = new ThemeProvider().getThemeForColor(props.themeMainColor);
        loadTheme({ ...{ palette: this._theme }, isInverted: false });

        this._selection = new Selection({
            onSelectionChanged: () => {
                this.setState({
                    selectedItemIds: this._getSelectedIds(),
                });

                if (this._selection.count === 1) {
                    this._previewSelectedFile();
                }
            },
        });

        this.state = {
            columns: this._getColumns(),
            gridItems: this._allItems,
            dataLoaded: false,
            filesUploading: false,
            selectedItemIds: [],
            sidePanelState: SidePanelState.closed,
            previewOpen: false,
            noteId: '',
            noteTitle: '',
            noteText: '',
            noteFile: null,
            previewedItem: null
        };

        this._loadData();
    }

    public render() {
        const { columns, gridItems } = this.state;

        const theme = getTheme();

        const columnProps: Partial<IStackProps> = {
            tokens: { childrenGap: 15 },
            styles: { root: { width: 300 } },
        };

        const buttonStyles = { root: { marginRight: 8 } };

        const closeButtonStyles: Partial<IButtonStyles> = {
            root: {
                color: theme.palette.neutralPrimary,
                marginLeft: 'auto',
                marginTop: '4px',
                marginRight: '2px',
            },
            rootHovered: {
                color: theme.palette.neutralDark,
            },
        };

        const prevButtonStyles: Partial<IButtonStyles> = {
            root: {
                color: theme.palette.neutralPrimary,
                //marginLeft: 'auto',
                marginTop: '15px',
                marginRight: '20px',
                marginBottom: '15px',
            },
            rootHovered: {
                color: theme.palette.neutralDark,
            },
        };

        const nextButtonStyles: Partial<IButtonStyles> = {
            root: {
                color: theme.palette.neutralPrimary,
                marginLeft: '20px',
                marginTop: '15px',
                //marginRight: 'auto',
                marginBottom: '15px',
            },
            rootHovered: {
                color: theme.palette.neutralDark,
            },
        };

        const labelStyles: Partial<ILabelStyles> = { root: { fontSize: 12, maxWidth: 150 } };

        const listStyles: Partial<IDetailsListStyles> = {
            contentWrapper: {
                color: '#333333',
                fontSize: 15,
                fontFamily: 'Segoe UI'
            },
            headerWrapper: {
                fontFamily: 'Segoe UI',
                fontSize: 12,
                fontWeight: 'normal',
                color: '#6e6f77',
                textAlign: 'left'
            }
        };

        const contentStyles = mergeStyleSets({
            container: {
                display: 'flex',
                flexFlow: 'column nowrap',
                alignItems: 'stretch',
                width: '70%',
                height: '80%',
                maxHeight: '80%',
                overflowY: 'hidden',
            },
            header: [
                theme.fonts.xLargePlus,
                {
                    flex: '1 1 auto',
                    borderTop: `4px solid ${theme.palette.themePrimary}`,
                    color: theme.palette.neutralPrimary,
                    display: 'flex',
                    alignItems: 'center',
                    fontWeight: FontWeights.semibold,
                    padding: '12px 12px 14px 24px',
                },
            ],
            body: {
                flex: '4 4 auto',
                padding: '0 24px 24px 24px',
                maxHeight: '80%',
                overflowY: 'scroll',
            },
        });

        const cmdItems: ICommandBarItemProps[] = [
            {
                key: 'newItem',
                text: this._props.context.resources.getString('attachNew'),
                iconProps: { iconName: 'Add' },
                onClick: this._openPanelCreateMode,
            },
            {
                key: 'refresh',
                text: this._props.context.resources.getString('refresh'),
                iconProps: { iconName: 'Refresh' },
                onClick: this._refreshData,
            },
            {
                key: 'delete',
                text: this._props.context.resources.getString('delete'),
                iconProps: { iconName: 'Delete' },
                disabled: this.state.selectedItemIds.length === 0,
                onClick: this._deleteSelectedNotes,
            },
        ];

        const fileButtonLabelName = this.state.sidePanelState === SidePanelState.create ? 'attachFile' : 'replaceFile';

        const cancelIcon: IIconProps = { iconName: 'Cancel' };

        return (
            <div>
                <div>
                    <CommandBar
                        items={[]}
                        farItems={cmdItems}
                    />
                </div>
                <div data-is-scrollable='true'>
                    <Panel
                        isOpen={this.state.sidePanelState !== SidePanelState.closed}
                        onDismiss={this._closeSidePanel}
                        headerText={this.state.sidePanelState === SidePanelState.create ? this._props.context.resources.getString('newNoteHeader') : this._props.context.resources.getString('editNoteHeader')}
                        onRenderFooterContent={this._renderPanelFooter}
                        isFooterAtBottom={true}
                    >
                        <Stack {...columnProps}>
                            <TextField value={this.state.noteTitle} onChange={this._onChangeNoteTitle} label={this._props.context.resources.getString('docTitle') + ' '} required />
                            <TextField value={this.state.noteText} onChange={this._onChangeNoteText} label={this._props.context.resources.getString('docDescription')} multiline autoAdjustHeight />
                            <Stack horizontal>
                                <PrimaryButton onClick={this._onAttachFile} styles={buttonStyles}>{this._props.context.resources.getString(fileButtonLabelName)}</PrimaryButton>
                                <Label styles={labelStyles}>{this.state.noteFile ? this.state.noteFile.fileName : ''}</Label>
                            </Stack>
                        </Stack>
                    </Panel>
                    <Stack horizontal>
                        <div id='subgrid_div'>
                            <MarqueeSelection selection={this._selection}>
                                <DetailsList
                                    items={gridItems}
                                    compact={false}
                                    columns={columns}
                                    selection={this._selection}
                                    selectionPreservedOnEmptyClick={true}
                                    getKey={this._getKey}
                                    setKey='multiple'
                                    layoutMode={DetailsListLayoutMode.justified}
                                    isHeaderVisible={true}
                                    styles={listStyles}
                                />
                            </MarqueeSelection>
                            {!this._allItems.length && (
                                <Stack horizontalAlign='center'>
                                    <Text>{this._props.context.resources.getString('noItems')}</Text>
                                </Stack>
                            )}
                            {!this.state.dataLoaded && (
                                <Stack horizontalAlign='center'>
                                    <Spinner size={SpinnerSize.large} label={this._props.context.resources.getString('loading')} />
                                </Stack>
                            )}
                        </div>
                        <div id='draganddrop_div' onDrop={this._fileDrop} onDragEnter={this._cancelDefaultBehavior} onDragOver={this._cancelDefaultBehavior} onDragLeave={this._cancelDefaultBehavior}>
                            {!this.state.filesUploading && (<p>{this._props.context.resources.getString('dropFiles')}</p>)}
                            {this.state.filesUploading && (
                                <Stack horizontalAlign='center'>
                                    <Spinner size={SpinnerSize.large} label={this._props.context.resources.getString('filesAreUploading')} />
                                </Stack>
                            )}
                        </div>
                    </Stack>
                </div>
                <Modal
                    titleAriaId={'previewDoc'}
                    allowTouchBodyScroll={false}
                    isOpen={this.state.previewOpen}
                    onDismiss={this._closePreview}
                    isBlocking={false}
                    containerClassName={contentStyles.container}
                >
                    <div className={contentStyles.header}>
                        <span id={'previewDocTitle'} style={{ fontSize: '1.5rem', margin: '10px auto auto 40px' }}>{this.state.previewedItem?.fileName}</span>
                        <IconButton onClick={this._closePreview} styles={closeButtonStyles} iconProps={cancelIcon} ariaLabel='Close' />
                    </div>
                    <Stack horizontal horizontalAlign='center'>
                        <DefaultButton onClick={this._previewPreviousFile} styles={prevButtonStyles} iconProps={{ iconName: 'Back' }}>{this._props.context.resources.getString('previousDocument')}</DefaultButton>
                        <DefaultButton onClick={this._previewNextFile} styles={nextButtonStyles} iconProps={{ iconName: 'Forward' }}>{this._props.context.resources.getString('nextDocument')}</DefaultButton>
                    </Stack>
                    <div className={contentStyles.body} style={{ display: this.state.previewedItem ? 'block' : 'none' }}>
                        <FileViewer fileType={this._getMimeTypeForPreview(true)} filePath={this._getFileForPreview()} />
                    </div>
                </Modal>
            </div>
        );
    }

    private _getKey(item: any): string {
        return item.Id;
    };

    private _getColumns = (): IColumn[] => {
        return [
            {
                key: 'column1',
                name: this._props.context.resources.getString('title'),
                fieldName: 'title',
                minWidth: 350,
                maxWidth: 500,
                isResizable: true,
                isSorted: true,
                isSortedDescending: false,
                sortAscendingAriaLabel: this._props.context.resources.getString('sortedAZ'),
                sortDescendingAriaLabel: this._props.context.resources.getString('sortedZA'),
                onColumnClick: this._onColumnClick,
                data: 'string',
                onRender: (item: INote) => {
                    return <a onClick={e => this._openPanelEditMode(item)}>{item.title}</a>;
                },
                isPadded: true,
            },
            {
                key: 'column2',
                name: this._props.context.resources.getString('description'),
                fieldName: 'description',
                minWidth: 350,
                maxWidth: 600,
                isResizable: true,
                isCollapsible: true,
                data: 'string',
                onColumnClick: this._onColumnClick,
                onRender: (item: INote) => {
                    return <span>{this._removeFormatting(item.description)}</span>;
                },
            },
            {
                key: 'column3',
                name: this._props.context.resources.getString('fileName'),
                fieldName: 'fileName',
                minWidth: 350,
                maxWidth: 500,
                isResizable: true,
                isSorted: false,
                isSortedDescending: false,
                sortAscendingAriaLabel: this._props.context.resources.getString('sortedAZ'),
                sortDescendingAriaLabel: this._props.context.resources.getString('sortedZA'),
                onColumnClick: this._onColumnClick,
                data: 'string',
                onRender: (item: INote) => {
                    return <a onClick={e => this._downloadAttachment(item.Id)}>{item.fileName}</a>;
                },
                isPadded: true,
            },
            {
                key: 'column4',
                name: this._props.context.resources.getString('createdOn'),
                fieldName: 'dateCreated',
                minWidth: 200,
                isResizable: true,
                isSorted: true,
                isSortedDescending: true,
                sortAscendingAriaLabel: this._props.context.resources.getString('sortedDateAcs'),
                sortDescendingAriaLabel: this._props.context.resources.getString('sortedDateDesc'),
                onColumnClick: this._onColumnClick,
                data: 'number',
                onRender: (item: INote) => {
                    return <span>{this.props.context.formatting.formatDateShort(item.dateCreated, true)}</span>;
                },
                isPadded: true,
            },
            //{
            //    key: 'column5',
            //    name: this._props.context.resources.getString('owner'),
            //    fieldName: 'owner',
            //    minWidth: 200,
            //    isResizable: true,
            //    isCollapsible: true,
            //    data: 'string',
            //    onColumnClick: this._onColumnClick,
            //    onRender: (item: INote) => {
            //        return <span>{item.owner}</span>;
            //    },
            //    isPadded: true,
            //},
        ];
    };

    private _loadData = (): void => {
        this._getNotes()
            .then(notes => {
                this._allItems = notes;

                this.setState({
                    gridItems: notes,
                    dataLoaded: true,
                });
            })
            .catch(err => {
                this._props.context.navigation.openErrorDialog({ message: this._props.context.resources.getString('errorNotesRetrieval') });
            });
    };

    private _refreshData = (): void => {
        this.setState({
            gridItems: [],
            dataLoaded: false
        });

        this._loadData();
    };

    private _renderPanelFooter = (): JSX.Element => {
        const buttonStyles = { root: { marginRight: 8 } };

        return (
            <div>
                <PrimaryButton onClick={this.state.sidePanelState === SidePanelState.create ? this._createNote : this._updateNote} styles={buttonStyles}>{this._props.context.resources.getString('save')}</PrimaryButton>
                <DefaultButton onClick={this._closeSidePanel}>{this._props.context.resources.getString('cancel')}</DefaultButton>
            </div>
        );
    };

    private _onChangeNoteTitle = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({
            noteTitle: newValue ? newValue : ''
        });
    };

    private _onChangeNoteText = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void => {
        this.setState({
            noteText: newValue ? newValue : ''
        });
    };

    private _closeSidePanel = (): void => {
        this.setState({
            sidePanelState: SidePanelState.closed,
            noteId: '',
            noteTitle: '',
            noteText: '',
            noteFile: null
        });
    };

    private _openPanelCreateMode = (): void => {
        this.setState({
            sidePanelState: SidePanelState.create
        });
    };

    private _openPanelEditMode = (item: INote): void => {
        this.setState({
            sidePanelState: SidePanelState.edit,
            noteId: item.Id,
            noteTitle: item.title,
            noteText: this._removeFormatting(item.description),
            noteFile: {
                fileContent: '',
                fileName: item.fileName,
                fileSize: 0,
                mimeType: ''
            }
        });
    };

    private _closePreview = (): void => {
        this.setState({
            previewOpen: false,
            previewedItem: null
        });
    };

    private _getSelectedIds = (): string[] => {
        return this._selection.getSelectedIndices().map((val) => { return this._allItems[val].Id });
    };

    private _removeFormatting = (text: string): string => {
        return sanitize(text, { allowedTags: [], allowedAttributes: {} }).trim();
    };

    private _onAttachFile = (): void => {
        this._props.context.device.pickFile()
            .then(files => {
                if (files && files.length > 0) {
                    this.setState({
                        noteFile: files[0]
                    });
                }
            })
            .catch(err => {
                this._props.context.navigation.openErrorDialog({ message: this._props.context.resources.getString('errorPickingFile') });
            });
    };

    private _createNote = (): void => {
        let newNote: any = {
            subject: this.state.noteTitle,
            notetext: this.state.noteText,
        };

        if (this.state.noteFile != null) {
            newNote.filename = this.state.noteFile.fileName;
            newNote.isdocument = true;
            newNote.documentbody = this.state.noteFile.fileContent;
            newNote.mimetype = this.state.noteFile.mimeType;
        }

        let entityId = this._props.entityId;
        let entityName = this._props.entityName;

        newNote['objectid_' + entityName + '@odata.bind'] = '/' + entityName + 's(' + entityId + ')';

        this._props.context.webAPI.createRecord('annotation', newNote).then(
            (result) => {
                this._closeSidePanel();

                this._refreshData();
            },
            (error) => {
                this._props.context.navigation.openErrorDialog({ message: this._props.context.resources.getString('errorNoteCreate') });
            }
        );
    };

    private _updateNote = (): void => {
        let noteUpdate: any = {
            subject: this.state.noteTitle,
            notetext: this.state.noteText,
        };

        if (this.state.noteFile != null && this.state.noteFile.fileContent != '') {
            noteUpdate.filename = this.state.noteFile.fileName;
            noteUpdate.isdocument = true;
            noteUpdate.documentbody = this.state.noteFile.fileContent;
            noteUpdate.mimetype = this.state.noteFile.mimeType;
        }

        this._props.context.webAPI.updateRecord('annotation', this.state.noteId, noteUpdate).then(
            (result) => {
                this._closeSidePanel();

                this._refreshData();
            },
            (error) => {
                this._props.context.navigation.openErrorDialog({ message: this._props.context.resources.getString('errorNoteUpdate') });
            }
        );
    };

    private _deleteSelectedNotes = (): void => {
        let confirmStrings: ComponentFramework.NavigationApi.ConfirmDialogStrings = {
            title: this._props.context.resources.getString('deleteConfirmationTitle'),
            text: this._props.context.resources.getString('deleteConfirmationText')

        };
        let confirmOptions: ComponentFramework.NavigationApi.ConfirmDialogOptions = {
            height: 200,
            width: 450
        };

        this._props.context.navigation.openConfirmDialog(confirmStrings, confirmOptions).then((res: ComponentFramework.NavigationApi.ConfirmDialogResponse) => {
            if (!res.confirmed) {
                return;
            }

            let deletePromises = this.state.selectedItemIds.map((noteId) => { return this._props.context.webAPI.deleteRecord('annotation', noteId); });

            Promise.all(deletePromises)
                .then((result) => {
                })
                .catch((error) => {
                    this._props.context.navigation.openErrorDialog({ message: this._props.context.resources.getString('errorNoteDelete') });
                })
                .finally(() => {
                    this._refreshData();
                });
        });
    };

    private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
        const { columns, gridItems } = this.state;
        const newColumns: IColumn[] = columns.slice();
        const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];

        newColumns.forEach((newCol: IColumn) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                currColumn.isSorted = true;
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });

        const newItems = this._copyAndSort(gridItems, currColumn.fieldName!, currColumn.isSortedDescending);
        this.setState({
            columns: newColumns,
            gridItems: newItems,
        });
    };

    private _copyAndSort<T>(items: T[], columnKey: string, isSortedDescending?: boolean): T[] {
        const key = columnKey as keyof T;
        return items.slice(0).sort((a: T, b: T) => ((isSortedDescending ? a[key] < b[key] : a[key] > b[key]) ? 1 : -1));
    };

    private _downloadAttachment = (id: string): void => {
        this._props.context.webAPI.retrieveRecord('annotation', id, '?$select=documentbody,filename,filesize,mimetype').then(
            result => {
                return {
                    fileContent: result['documentbody'],
                    fileName: result['filename'],
                    fileSize: result['filesize'],
                    mimeType: result['mimetype']
                };
            }).then(fileObj => {
                this._props.context.navigation.openFile(fileObj, { openMode: 2 });
            });
    };

    private _getNotes = (): Promise<INote[]> => {
        let recordId = (this.props.context as any).page.entityId;
        let query = '?$select=annotationid,subject,notetext,createdon,filename,_ownerid_value&$expand=owninguser($select=fullname),owningteam($select=name)&$filter=(_objectid_value eq ' + recordId + ' and filename ne null)&$orderby=createdon desc,subject asc';

        return this._props.context.webAPI.retrieveMultipleRecords('annotation', query)
            .then(result => {
                let notes: INote[] = [];
                for (let i = 0; i < result.entities.length; i++) {
                    let ent = result.entities[i];

                    notes.push({
                        Id: ent['annotationid'].toString(),
                        title: ent['subject'] ? ent['subject'].toString() : '',
                        description: ent['notetext'] ? ent['notetext'].toString() : '',
                        fileName: ent['filename'].toString(),
                        owner: ent['owninguser'] ? ent['owninguser']['fullname'] : ent['owningteam']['name'],
                        dateCreated: new Date(ent['createdon'].toString()),
                    });
                }

                return notes;
            })
            .catch(err => {
                this._props.context.navigation.openErrorDialog({ message: this._props.context.resources.getString('errorNotesRetrieval') });
                return [];
            });
    };

    private _cancelDefaultBehavior = (e: React.DragEvent<HTMLDivElement>): void => {
        e.preventDefault();
        e.stopPropagation();
    };

    private _fileDrop = (e: React.DragEvent<HTMLDivElement>): void => {
        if (this.state.filesUploading) {
            this._props.context.navigation.openAlertDialog({ text: this._props.context.resources.getString('uploadInProgress') });
            return;
        }

        e.preventDefault();

        let files: FileList = e.dataTransfer.files;
        if (!files || files.length < 1) {
            this._props.context.navigation.openAlertDialog({ text: this._props.context.resources.getString('noFilesAttached') });
            return;
        }

        this.setState({ filesUploading: true });

        for (let i = 0; i < files.length; i++) {
            this._uploadFile(files[i]);
        }

        this._interval = setInterval(this._handleUploadCompletion, 500);
    };

    private _uploadFile = (file: File): void => {
        this._filesUpload.files.push(file);
        const reader: FileReader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => {
            let fileString: string = reader.result ? reader.result.toString() : '';
            let note: any = {
                subject: file.name,
                filename: file.name,
                isdocument: true,
                documentbody: fileString.length > 0 ? fileString.substring(fileString.indexOf(',') + 1, fileString.length) : '',
                filesize: file.size,
                mimetype: file.type,
                objecttypecode: this._props.entityName,
                ['objectid_' + this._props.entityName + '@odata.bind']: this._props.entityName + 's(' + this._props.entityId + ')'
            };

            this._props.context.webAPI.createRecord('annotation', note)
                .then(result => {
                    this._filesUpload.createdNotesRefs.push(result);
                })
                .catch(err => {
                    this._filesUpload.errors.push(err.message.toLowerCase());
                })
        };
    };

    private _handleUploadCompletion = (): void => {
        if (this._filesUpload.createdNotesRefs.length + this._filesUpload.errors.length < this._filesUpload.files.length) {
            return;
        }

        clearInterval(this._interval);

        if (this._filesUpload.errors.length > 0) {
            const fileTooBig: boolean = this._filesUpload.errors.some((errMsg) => { return errMsg.includes('exceeds the maximum size limit'); });
            this._props.context.navigation.openErrorDialog({ message: this._props.context.resources.getString(fileTooBig ? 'errorFileTooBig' : 'errorFilesUpload') });
        }

        this._filesUpload = {
            files: [],
            createdNotesRefs: [],
            errors: []
        };

        this._refreshData();

        this.setState({ filesUploading: false });
    };

    private _previewSelectedFile = (): void => {
        let selectedNote = this._allItems.find((item: INote) => { return item.Id === this._getSelectedIds()[0]; });

        this._previewDocument(selectedNote);
    };

    private _previewNextFile = (): void => {
        let nextNote = this._getPreviousOrNextDocument(true);

        this._previewDocument(nextNote);
    };

    private _previewPreviousFile = (): void => {
        let previousNote = this._getPreviousOrNextDocument(false);

        this._previewDocument(previousNote);
    };

    private _previewDocument = (note: INote | undefined): void => {
        this.setState({
            previewedItem: null
        });

        if (!note) {
            return;
        }

        if (note.file) {
            setTimeout(() => {
                this.setState({
                    previewOpen: true,
                    previewedItem: note
                });
            }, 100);
        }
        else {
            this._props.context.webAPI
                .retrieveRecord('annotation', note.Id, '?$select=annotationid,documentbody,filename,filesize,mimetype')
                .then((result) => {
                    note.file = {
                        fileName: result['filename'],
                        fileSize: result['filename'],
                        mimeType: result['mimetype'],
                        fileContent: result['documentbody']
                    };

                    this.setState({
                        previewOpen: true,
                        previewedItem: note
                    });
                });
        }
    };

    private _getPreviousOrNextDocument = (getNext: boolean): INote | undefined => {
        const noteId = this.state.previewedItem?.Id;
        if (!noteId) {
            return;
        }

        let noteIndex: number = this._allItems.findIndex((item: INote) => { return item.Id === noteId; });
        let newIndex: number;
        if (getNext) {
            newIndex = noteIndex < this._allItems.length - 1 ? noteIndex + 1 : 0;
        }
        else {
            newIndex = noteIndex > 0 ? noteIndex - 1 : this._allItems.length - 1;
        }

        return this._allItems[newIndex];
    };

    private _getMimeTypeForPreview = (truncate: boolean): string => {
        let mimeType = this.state.previewedItem?.file?.mimeType;
        if (!mimeType) {
            return '';
        }

        switch (mimeType) {
            case 'application/msword':
            case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            case 'application/octet-stream': // plugin version of document generator incorrectly sets mime type for generated word documents as octet-stream
                mimeType = 'application/docx';
                break;
            case 'application/vnd.ms-excel':
                mimeType = 'application/csv';
                break;
            case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                mimeType = 'application/xslx';
                break;
        }

        return truncate ? mimeType.substring((mimeType.indexOf('/') + 1), mimeType.length) : mimeType;
    };

    private _getFileForPreview = (): string => {
        return 'data:' + this._getMimeTypeForPreview(false) + ';base64, ' + this.state.previewedItem?.file?.fileContent;
    };


}
