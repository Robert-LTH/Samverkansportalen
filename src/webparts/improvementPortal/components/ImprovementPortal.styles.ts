import { mergeStyleSets } from 'office-ui-fabric-react/lib/Styling';

const styles = mergeStyleSets({
  portal: {
    padding: 20
  },
  header: {
    marginBottom: 16
  },
  controls: {
    marginBottom: 16
  },
  voteSummary: {
    marginBottom: 20
  },
  form: {
    backgroundColor: '#faf9f8',
    border: '1px solid #edebe9',
    borderRadius: 4,
    padding: 16,
    marginBottom: 24
  },
  formButtons: {
    marginTop: 12,
    display: 'flex',
    gap: 8
  },
  searchRow: {
    display: 'flex',
    flexWrap: 'wrap',
    gap: 12,
    alignItems: 'flex-end',
    marginBottom: 16
  },
  sectionTitle: {
    margin: '24px 0 8px',
    fontWeight: 600
  },
  suggestionCard: {
    background: '#fff',
    border: '1px solid #edebe9',
    borderRadius: 6,
    padding: 16,
    marginBottom: 12,
    boxShadow: '0 1px 3px rgba(0, 0, 0, 0.08)'
  },
  cardHeader: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    gap: 12
  },
  statusBadge: {
    background: '#f3f2f1',
    color: '#323130',
    padding: '4px 10px',
    borderRadius: 20,
    fontSize: 12,
    fontWeight: 600
  },
  meta: {
    color: '#605e5c',
    fontSize: 12,
    marginTop: 4,
    marginBottom: 12
  },
  description: {
    whiteSpace: 'pre-line',
    marginBottom: 16
  },
  voteActions: {
    display: 'flex',
    alignItems: 'center',
    gap: 12,
    flexWrap: 'wrap'
  },
  voteCounter: {
    fontWeight: 600
  },
  statusControls: {
    display: 'flex',
    alignItems: 'center',
    gap: 8,
    flexWrap: 'wrap',
    marginTop: 16
  },
  emptyState: {
    padding: '32px 16px',
    textAlign: 'center',
    color: '#8a8886'
  },
  messageBar: {
    marginBottom: 12
  }
});

export type ImprovementPortalStyles = typeof styles;

export default styles;
